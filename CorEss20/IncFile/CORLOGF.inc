<%

' ****************************************************************************************
' ** Cor E-Solution v48 -    Esquema de Configuracion ------------------ marzo, 2020     **
' ****************************************************************************************
'Dim STR,STRF

Dim sub1
''XLCHARSET = "iso-8859-1"
Const ORDERTYPE_ASC   = 10
Const ORDERTYPE_DESC  = 11
Const SORTBY_KEY      = 20
Const SORTBY_VALUE    = 21
Dim AUsrG2
XLCHARSET = "windows-1252"
'XLCHARSET = "utf-8"

XLFECDAT   = 0
XLHORDAT   = 10
XLFECDAT2  = 0
XLHORDAT2  = 10
XLLOGID    = 0
XLLOGIDDET = 0
XLLOGIDEXT = 0

GLBSH      = 0
GLBSM      = 0
GLBSS      = 0

TAEMP      = 0
IAEMP      = 0
Dim ATREMP(50,5)
Dim cmp

USREMP = 1

Dim ATXCOM(20,4)
GLBCOMBO    = 0
GLBCACTU    = 0


Dim ATABGLOSLNKS(500,2)
TTLNKS = 0


XGUTABLISKEY  = ""
XGUTABLISKEY2 = ""
XGUTABLISDAT  = ""

XGUMASTER = "0"
XGUFCOBJ  = "0"
XGUFTOBJ  = "0"
XGUFNIV1  = "0"
XGUFNIV2  = "0"
XGUFNIV3  = "0"
XGUFNIV4  = "0"
XGUFNIV5  = "0"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  Sets de configuraciones de Panel
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


XGBinput1   = ""
XGBinput2   = ""
XGBinput3   = ""
XGBinput4   = ""
XGBinput5   = ""
XGBinput6   = ""
XGBinput7   = ""


Styl    = ""
Sglos2  = ""
Sglos1  = ""
Sglos3  = ""

SfidkA  = ""
SfidkC  = ""
SfidkB  = ""

Sfidk0  = ""
Sfidk1  = ""
Sfidk2  = ""
Sfidk3  = ""
Sfidk4  = ""

Sglcmp = ""
Sglcnt = ""
Sglprf = ""
Sgllnk = ""


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  Sets de configuraciones generales
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' Aplicacion
' por los nuevos menus en el frame
'tiene que venir como parametro cuando arma los menus
Ecoraplic=request.querystring("coraplic")

if Ecoraplic <> "" then Session("CorApli")=Ecoraplic

If Session("CorApli") = "COREM20" Then
	wkaplic = "EM"
ElseIf Session("CorApli") = "COREP20" Then
	wkaplic = "PL"
ElseIf Session("CorApli") = "CLI916953113116518847748" Then
	wkaplic = "GP"
ElseIf Session("CorApli") = "CORESS20" Then
	wkaplic = "CE"
ElseIf Session("CorApli") = "CLI111557214260398409701" Then
	wkaplic = "OR"
Else
	wkaplic = "RG"
End If




If session("Usuario") = "" Then
	session("LUID") = ofv.ADbUid
	session("LPWD") = ofv.ADbPwd
End If

ofv.ADbUidl = session("LUID")
ofv.ADbPwdl = session("LPWD")


If request.querystring("MNN") <> "" Then
	session("FormN") = Replace(request.querystring("MNN"),"_"," ")
End If

If request.querystring("seg") <> "" Then
	session("Form") = Replace(request.querystring("seg"),"_"," ")
End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Funciones de administracion de Windows
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GTkey(XCodobJ)

	GThms   = DateDiff("d",Date(),CDate("12/12/2050")) & FormatDateTime(Now(),3)
	GTses   = session.sessionid
	GTsecu  = CStr(XCodobJ & GThms & GTses)

	GTsecuf = ""

	For GTi=1 To Len(GTsecu)
		GTcs = Mid(GTsecu,GTi,1)
		If  GTcs => "0" And GTcs <= "9" Then
			GTsecuf = GTsecuf & GTcs
		End If
	Next

	GTkey  = Trim(GTsecuf)

End Function

Function ConsCambios(sASP)

	Dim wxs
	wxs = wxs & "<script language=JavaScript>" & Chr(13)
	wxs = wxs & " function AbrirCambios()" & Chr(13)
	wxs = wxs & " {" & Chr(13)
	If BrowserName() = "IE" Then
		wxs = wxs & " window.showModalDialog(" & Chr(34) & sASP & Chr(34) & ",window,"
		wxs = wxs &  Chr(34) & "dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"
		wxs = wxs & "dialogTop:250px;dialogLeft:300px;status:no" & Chr(34)
	Else
		wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34) & ",window,"
		wxs = wxs &  Chr(34) & "status=no,Width=600,Height=350,top=250,left=300"
		wxs = wxs & Chr(34)
	End If


	wxs = wxs & " ); " & Chr(13)
	wxs = wxs & "}" & Chr(13)
	wxs = wxs & " </script>" & Chr(13)

	ConsCambios =      wxs

End Function

Function Ordenar(sASP)

	Dim wxs

	wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bor1", "Ordenar", "button", "", "10", "bt", "AbrirOrdenar()") & "<br><br>", "si")
	wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

	wxs = wxs & "<script language=JavaScript>" & Chr(13)
	wxs = wxs & "function AbrirOrdenar() { " & Chr(13)
	wxs = wxs & " var h=570;w=790 ;"
	If  session("APLMDCLI")="BC" Then
		wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If
	If BrowserName()="IE" Then
		wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
		'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
	Else
		wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
	End if'
	wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
	wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
	wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
	wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

	Ordenar =      wxs

End Function

Function Aprocesos(sASP)



	Dim wxs

	wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bor1", "Procesos", "button", "", "10", "bt", "AbrirProcesos()") & "<br><br>", "si")
	wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

	wxs = wxs & "<script language=JavaScript>" & Chr(13)
	wxs = wxs & "function AbrirProcesos() { " & Chr(13)
	wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
	wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
	wxs = wxs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
	wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
	wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

	Aprocesos =      wxs

End Function


Function Editar(sASP)



	Dim wxs

	wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bed1", "Editar", "button", "", "10", "bt", "AbrirVentana()") & "<br><br>", "si")
	wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

	wxs = wxs & "<script language=JavaScript>" & Chr(13)
	wxs = wxs & "function AbrirVentana() {" & Chr(13)
	wxs = wxs & " var h=570;w=790 ;"
	'	If  session("APLMDCLI")="BC" Then
	wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'	End If
	If BrowserName()="IE" Then
		wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
		'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
	Else
		wxs = wxs & "CorWin=window.opener.open(" & Chr(34) & sASP & Chr(34)
	End If
	wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
	wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
	wxs = wxs & " ); " & Chr(13) & "window.close();" & Chr(13)
	wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

	Editar =      wxs

End Function

Function Eliminar(sASP)


	Dim      wxs

	wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bel1", "Eliminar", "button", "", "10", "bt", "AbrirEliminar()") & "<br><br>", "si")
	wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

	wxs = wxs & "<script language=JavaScript> " & Chr(13)
	wxs = wxs & "function AbrirEliminar() {" & Chr(13)
	'wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
	wxs = wxs & " var h=570;w=790 ;"
	'If  session("APLMDCLI")="BC" Then
		wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'End If
	If BrowserName()="IE" Then
		wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
		'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
	Else
		wxs = wxs & "window.opener.open(" & Chr(34) & sASP & Chr(34)
	End if'
	wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
	wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
	wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
	wxs = wxs & " } " & Chr(13) & " </script>" & Chr(13)

	Eliminar =      wxs


End Function

Function Copiar(sASP)



	Dim      wxs

	wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bco1", "Copiar", "button", "", "10", "bt", "AbrirCopiar()") & "<br><br>", "si")
	wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

	wxs = wxs & "<script language=JavaScript>" & Chr(13)
	wxs = wxs & "function AbrirCopiar() {" & Chr(13)
	wxs = wxs & " var h=570;w=790 ;"
	'If  session("APLMDCLI")="BC" Then
		wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'End If
	If BrowserName()="IE" Then
		wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
		'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
	Else
		wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
	End if'
	wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
	wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
	wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
	wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

	Copiar =      wxs


End Function

Function MOVER(sASP)

	Dim wxs

	wxs = ofv.GenCelda("", "", "center", "33%", "", "", "",   ofv.GenerarInput("bco1", "Mover", "button", "", "10", "bt", "AbrirMover()") & "<br><br>"  , "si")
	wxs = ofv.GenForm("", "", "", "", "post", "", wxs, "si")

	wxs = wxs & "<script language=JavaScript>" & Chr(13)


	wxs = wxs & "function AbrirMover() {" & Chr(13)
	wxs = wxs & " var h=570;w=790 ;"
'	If  session("APLMDCLI")="BC" Then
		wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
'	End If
	If BrowserName()="IE" Then
		wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
		'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
	Else
		wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
	End if'
	wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
	wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
	wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
	wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

	MOVER = wxs

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Funciones de administracion de EMail
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SendMailAtt(ByVal strEmailBodyMessage, ByVal strRecipientEmailAddress , ByVal strFromEmailAddress, ByVal strSubject, strMailComponent, blnHTML,ByVal Importance, ByVal notification, ByVal att, ByVal mailfijo,ByVal tipoenv)

	'Dimension variables
	Dim objCDOSYSMail		'Holds the CDOSYS mail object
	Dim objCDOMail			'Holds the CDONTS mail object
	Dim objJMail			'Holds the Jmail object
	Dim objAspEmail			'Holds the Persits AspEmail email object
	Dim objAspMail			'Holds the Server Objects AspMail email object
	Dim strEmailBodyAppendMessage	'Holds the appended email message
	Dim SmtpServer
	Dim SmtpPort

	SmtpServer = Session("SMTPServer")
	SmtpPort   = Session("SMTPPort")

	' RESPONSE.WRITE strRecipientEmailAddress & "<BR>" & strFromEmailAddress
	SHTMLMail = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
	SHTMLMail = SHTMLMail & "<html>"
	SHTMLMail = SHTMLMail & "<head>"
	SHTMLMail = SHTMLMail & "<title>" & subject & "</title>"
	SHTMLMail = SHTMLMail & "</head>"
	SHTMLMail = SHTMLMail & "<body bgcolor=""FFFFFF"">"
	SHTMLMail = SHTMLMail & strEmailBodyMessage

	SHTMLMail = SHTMLMail & "</body>"
	SHTMLMail = SHTMLMail & "</html>"

	'Dimension variables
	Dim objCDOSYSCon

	'Create the e-mail server object
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
	objCDOSYSMail.HTMLBody = SHTMLMail
	'Set and update fields properties

	'Out going SMTP server
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SmtpServer
	'SMTP port
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  =SmtpPort
	'CDO Port
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

	' esto es por google no va en los clientes ---------------------------------------
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "soptec@corconsulting.com.ar"
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Pichon1612"
	'-------------------------------------------------------------------------------------


	'Timeout
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	objCDOSYSCon.Fields.Update


	'Update the CDOSYS Configuration
	Set objCDOSYSMail.Configuration = objCDOSYSCon


	'Who the e-mail is from
	'objCDOSYSMail.From = strFromEmailName & " <" & strFromEmailAddress & ">"
	objCDOSYSMail.From = cstr(strFromEmailAddress & "")

	'Who the e-mail is sent to
	'objCDOSYSMail.To = strRecipientName & " <" & strRecipientEmailAddress & ">"
	objCDOSYSMail.To =  cstr( strRecipientEmailAddress & "")

	objCDOSYSMail.CC = cstr(mailfijo & "")

	'The subject of the e-mail
	objCDOSYSMail.Subject = strSubject

	'Set the e-mail body format (HTMLBody=HTML TextBody=Plain)
	If UCase(blnHTML)="SI" Then
		objCDOSYSMail.HTMLBody = SHTMLMail
	Else
		objCDOSYSMail.TextBody = SHTMLMail
	End If
	If UCase(notification)="SI" Then
		objCDOSYSMail.fields("urn:schemas:mailheader:disposition-notification-to") =chr(34) & strFromEmailAddress & Chr(34) ' "<valid e-mail address>" 'ToDo: Type a valid e-mail address.
		objCDOSYSMail.fields("urn:schemas:mailheader:return-receipt-to") = chr(34) & strFromEmailAddress & Chr(34) 			' "<valid e-mail address>"  'ToDo: Type a valid e-mail address.
	End If
	Select Case importance
		Case 0 'BAJA
		objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 5
		Case 1 'NORMAL
		objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 3
		Case 2  'ALTA
		objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 1

	End Select
	' response.write att
	If tipoenv = 1 Then
		objCDOSYSMail.AddAttachment (CStr(att & ""))
	End If

	objCDOSYSMail.Fields.Update

	'Send the e-mail
	On Error Resume Next

	objCDOSYSMail.Send


	'Close the server mail object
	Set objCDOSYSMail = Nothing


End Function

Function SendMail(ByVal strEmailBodyMessage, ByVal strRecipientEmailAddress , ByVal strFromEmailAddress, ByVal strSubject, strMailComponent, blnHTML,ByVal Importance, ByVal notification)

	'Dimension variables
	Dim objCDOSYSMail		'Holds the CDOSYS mail object
	Dim objCDOMail			'Holds the CDONTS mail object
	Dim objJMail			'Holds the Jmail object
	Dim objAspEmail			'Holds the Persits AspEmail email object
	Dim objAspMail			'Holds the Server Objects AspMail email object
	Dim strEmailBodyAppendMessage	'Holds the appended email message
	Dim SmtpServer
	Dim SmtpPort

	SmtpServer = Session("SMTPServer")
	SmtpPort   = Session("SMTPPort")

	' RESPONSE.WRITE strRecipientEmailAddress & "<BR>" & strFromEmailAddress
	SHTMLMail = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
	SHTMLMail = SHTMLMail & "<html>"
	SHTMLMail = SHTMLMail & "<head>"
	SHTMLMail = SHTMLMail & "<title>" & subject & "</title>"
	SHTMLMail = SHTMLMail & "</head>"
	SHTMLMail = SHTMLMail & "<body bgcolor=""FFFFFF"">"
	SHTMLMail = SHTMLMail & strEmailBodyMessage

	SHTMLMail = SHTMLMail & "</body>"
	SHTMLMail = SHTMLMail & "</html>"


	'******************************************
	'***	        Mail components        ****
	'******************************************

	'Select which email component to use
	Select Case UCase(strMailComponent)


		'******************************************
		'***	  MS no envia     ****
		'******************************************

		'CDOSYS mail component
		Case "NOSEND"



		'******************************************
		'***	  MS CDOSYS mail component     ****
		'******************************************

		'CDOSYS mail component
		Case "CDOSYS"

		'Dimension variables
		Dim objCDOSYSCon

		'Create the e-mail server object
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

		'Set and update fields properties

		'Out going SMTP server
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SmtpServer
		'SMTP port
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  =SmtpPort
		'CDO Port
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2



		'Timeout
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		objCDOSYSCon.Fields.Update


		'Update the CDOSYS Configuration
		Set objCDOSYSMail.Configuration = objCDOSYSCon


		'Who the e-mail is from
		'objCDOSYSMail.From = strFromEmailName & " <" & strFromEmailAddress & ">"
		objCDOSYSMail.From = strFromEmailAddress

		'Who the e-mail is sent to
		'objCDOSYSMail.To = strRecipientName & " <" & strRecipientEmailAddress & ">"
		objCDOSYSMail.To =  strRecipientEmailAddress
		If OPWK = "REASIGNAR" Then
			objCDOSYSMail.CC = usrcc
		End If
		'The subject of the e-mail
		objCDOSYSMail.Subject = strSubject

		'Set the e-mail body format (HTMLBody=HTML TextBody=Plain)
		If UCase(blnHTML)="SI" Then
			objCDOSYSMail.HTMLBody = SHTMLMail
		Else
			objCDOSYSMail.TextBody = SHTMLMail
		End If
		If UCase(notification)="SI" Then
			objCDOSYSMail.fields("urn:schemas:mailheader:disposition-notification-to") =chr34 & strFromEmailAddress & Chr(34) ' "<valid e-mail address>" 'ToDo: Type a valid e-mail address.
			objCDOSYSMail.fields("urn:schemas:mailheader:return-receipt-to") = chr34 & strFromEmailAddress & Chr(34) 			' "<valid e-mail address>"  'ToDo: Type a valid e-mail address.
		End If
		Select Case importance
			Case 0 'BAJA
			objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 5
			Case 1 'NORMAL
			objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 3
			Case 2  'ALTA
			objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 1

		End Select

		objCDOSYSMail.Fields.Update

		'Send the e-mail
		On Error Resume Next
		objCDOSYSMail.Send


		'Close the server mail object
		Set objCDOSYSMail = Nothing




		'******************************************
		'***  	  MS CDONTS mail component     ****
		'******************************************

		'CDONTS mail component
		Case "CDONTS"






		'Create the e-mail server object
		Set objCDOMail = Server.CreateObject("CDONTS.NewMail")


		'Who the e-mail is from
		objCDOMail.From =  strFromEmailAddress

		'Who the e-mail is sent to
		objCDOMail.To = strRecipientEmailAddress

		'The subject of the e-mail
		objCDOMail.Subject = strSubject

		'The main body of the e-amil
		objCDOMail.Body = SHTMLMail

		'Set the e-mail body format (0=HTML 1=Text)
		If blnHTML ="SI" Then
			objCDOMail.BodyFormat = 0
		Else
			objCDOMail.BodyFormat = 1
		End If

		'Set the mail format (0=MIME 1=Text)
		objCDOMail.MailFormat = 0

		'Importance of the e-mail (0=Low, 1=Normal, 2=High)
		objCDOMail.Importance = importance

		If UCase(notification) = "SI" Then
			strReply_To = "<" & From & ">"
			objCDOMail.Value("Disposition-Notification-To") = strReply_To
		End If

		'Send the e-mail
		objCDOMail.Send


		'Close the server mail object
		Set objCDOMail = Nothing




		'******************************************
		'***  	  w3 JMail mail component      ****
		'******************************************

		'JMail component
		Case "Jmail"

		'Create the e-mail server object
		Set objJMail = Server.CreateObject("JMail.SMTPMail")

		With objJMail
			'Out going SMTP mail server address
			.ServerAddress = strIncomingMailServer

			'Who the e-mail is from
			.Sender = strFromEmailAddress
			.SenderName = strFromEmailName

			'Who the e-mail is sent to
			.AddRecipient strRecipientEmailAddress

			'The subject of the e-mail
			.Subject = strSubject

			'Set the e-mail body format (BodyHTML=HTML Body=Text)
			If blnHTML = True Then
				.HTMLBody = strEmailBodyMessage & strEmailBodyAppendMessage
			Else
				.Body = strEmailBodyMessage & strEmailBodyAppendMessage
			End If

			'Importance of the e-mail
			.Priority = 3

			'Send the e-mail
			If Not strIncomingMailServer = "" Then .Execute
		End With

		'Close the server mail object
		Set objJMail = Nothing




		'******************************************
		'*** Persits AspEmail mail component   ****
		'******************************************

		'AspEmail component
		Case UCase("AspEmail")
		response.write strMailComponent
		'Create the e-mail server object
		Set objAspEmail = Server.CreateObject("Persits.MailSender")


		'Out going SMTP mail server address
		objAspEmail.Host = SmtpServer

		'Who the e-mail is from
		objAspEmail.From = strFromEmailAddress
		'.FromName = strFromEmailName

		'Who the e-mail is sent to
		objAspEmail.AddAddress strRecipientEmailAddress
		'  response.write strRecipientEmailAddress
		'The subject of the e-mail
		objAspEmail.Subject = strSubject

		'Set the e-mail body format (BodyHTML=HTML Body=Text)
		If UCase(blnHTML) = "SI" Then
			objAspEmail.IsHTML = True
		End If

		'The main body of the e-mail
		objAspEmail.Body = SHTMLMail

		Select Case importance
			Case 0 'BAJA
			objAspEmail.Priority=5
			Case 1 'NORMAL
			objAspEmail.Priority=3
			Case 2  'ALTA
			objAspEmail.Priority=1
		End Select


		If UCase(notification)="SI" Then
			objAspEmail.AddCustomHeader "Return-Receipt-To: " & Chr(34) & strFromEmailAddress & Chr(34)
		End If


		'Send the e-mail
		objAspEmail.Send


		'Close the server mail object
		Set objAspEmail = Nothing




		'********************************************
		'*** ServerObjects AspMail mail component ***
		'********************************************

		'AspMail component
		Case "AspMail"

		'Create the e-mail server object
		Set objAspMail = Server.CreateObject("SMTPsvg.Mailer")

		With objAspMail
			'Out going SMTP mail server address
			.RemoteHost = strIncomingMailServer

			'Who the e-mail is from
			.FromAddress = strFromEmailAddress
			.FromName = strFromEmailName

			'Who the e-mail is sent to
			.AddRecipient " ", strRecipientEmailAddress

			'The subject of the e-mail
			.Subject = strSubject

			'Set the e-mail body format (BodyHTML=HTML Body=Text)
			If blnHTML = True Then
				.ContentType = "text/HTML"
			End If

			'The main body of the e-mail
			.BodyText = strEmailBodyMessage & strEmailBodyAppendMessage

			'Send the e-mail
			If Not strIncomingMailServer = "" Then .SendMail
		End With

		'Close the server mail object
		Set objAspMail = Nothing
	End Select

	'Set the returned value of the function to true
	SendMail = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' COR LOGF Include -  Sets de config. general y funciones - V 2.8.6   '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''' funciones del glosario ''''''''''''''''''''''''''''

Function GenKeyLnks(gtlnks)

	If session("usuario")="GOOGLE" Then
		GenKeyLnks="(XLNK:XLNK)"
	Else
		GenKeyLnks = Trim(gtlnks)
	End If

End Function



Function GenKeyLnks2(gtlnks)

	TTLNKS = TTLNKS + 1

	sh = DatePart("h",Time())
	sm = DatePart("n",Time())
	ss = DatePart("s",Time())

	XCodobJ = sh & sm & ss

	GThms   = DateDiff("d",Date(),CDate("12/12/2050")) & FormatDateTime(Now(),3)
	GTses   = session.sessionid
	GTsecu  = CStr(XCodobJ & GThms & GTses)

	GTsecuf = "(XLNK:" & TTLNKS & ":"


	For GTi=1 To Len(GTsecu)
		GTcs = Mid(GTsecu,GTi,1)
		If GTcs => "0" And GTcs <= "9" Then
			GTsecuf = GTsecuf & GTcs
		End If
	Next


	GTsecuf = GTsecuf & "XLNK)"

	ATABGLOSLNKS(TTLNKS,1) = Trim(GTsecuf)
	ATABGLOSLNKS(TTLNKS,2) = gtlnks

	GenKeyLnks2 = Trim(GTsecuf)

End Function



Function Capitalize (str)
	Capitalize=UCase(Left(str, 1)) & LCase(Mid(str, 2))
End Function



Function VerGlosario(XEhtml)

	VerGlosario = XEhtml

End Function

Function armaglos()
	Dim  RSX1zz
	cadena=""
	strsqlP = "Select KEYWORD,SHWWORD From GLOSARIO ORDER BY SHWWORD"
	Set cccn = ofv.conectar(ofv.strconn2)

	Set rsX1zz = ofv.crearconsultaEx(StrSqlP,cccn,1,parametros)
	i=0

	Do While Not RSX1zz.EOF

		cadena= cadena & "Aglos[" & i &"]='" & rsx1zz(0) &"';" & Chr(13)
		rsX1zz.movenext
		i=i+1
	Loop
	armaglos=cadena

	If session("yider")<> "" Then
		cadena=""
		session("yider")=Replace(session("yider"),"*" ,"")
		ayider=Split(session("yider")," ")
		If IsArray(ayider) Then
			zdelim= ""
			pals=""
			'i=0
			For ia=0 To  -1 'ubound(ayider)
				pal= Trim(ayider(ia) & "")
				If Len(pal) > 3 Then
					cadena= cadena & "Aglos[" & i &"]='" & pal &"';" & Chr(13)
					i=I+1
				End If
				' response.write pal
			Next
			armaglos=armaglos & cadena
		End If

		Call OFV.CERRARCONSULTA (rsX1zz)
		Call OFV.CERRARCONN(cccn)
	End If
	' response.write session("yider")
End Function

Function VerGlosario2(XEhtml)

	Dim ATABGLOS(200,5)
	ttglos = 0

	XEconvMl = "<span CLASS=gls onclick=" & Chr(34) & "verglos2('XXARG')" & Chr(34) & " >XXDATA</span>"
	' XEconvMl = "XXDATA"

	' glosario

	strsql = "Select KEYWORD,SHWWORD From GLOSARIO ORDER BY SHWWORD"
	Set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rsX1.eof Then
		Do While Not RSX1.EOF

			TTGLOS = TTGLOS + 1
			ATABGLOS(TTGLOS,0) = RSX1(0)

			ATABGLOS(TTGLOS,1) = RSX1(1)

			If  UCase(ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) Then
				ATABGLOS(TTGLOS,2) = "XXXXXX"
			Else
				ATABGLOS(TTGLOS,2) = UCase(ATABGLOS(TTGLOS,1))
			End If

			If  LCase(ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) Then
				ATABGLOS(TTGLOS,3) = "XXXXXX"
			Else
				ATABGLOS(TTGLOS,3) = LCase(ATABGLOS(TTGLOS,1))
			End If

			If  Capitalize (ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) Then
				ATABGLOS(TTGLOS,4) = "XXXXXX"
			Else
				ATABGLOS(TTGLOS,4) = Capitalize (ATABGLOS(TTGLOS,1))
			End If





			RSX1.MOVENEXT
		Loop

	End If
	Call ofv.cerrarconsulta(rsX1)

	If TTGLOS > 0 Then

		For IGLOS = 1 To TTGLOS
			XEconvBs = Replace(XEconvML,"XXARG","XXARG" & (999-IGLOS))
			For III = 1 To 4
				If  ATABGLOS(IGLOS,III) <> "XXXXXX" Then
					XEconv = Replace(XEconvBs,"XXDATA",ATABGLOS(IGLOS,III))

					XEhtml = Replace(XEhtml, ATABGLOS(IGLOS,III),XEconv)
				End If
			Next
		Next

		For IGLOS = 1 To TTGLOS
			XEhtml   = Replace(XEhtml,"XXARG" & (999-IGLOS),ATABGLOS(IGLOS,0))
		Next

	End If

	For IGLOS = 1 To TTLNKS
		XEhtml   = Replace(XEhtml, ATABGLOSLNKS(IGLOS,1),ATABGLOSLNKS(IGLOS,2))
	Next

	xehtmln = InStr(XEhtml,"<span CLASS=gls")

	xehtmlx = ""


	If xehtmln = 0 Then

		XEhtml = Replace(XEhtml,"versup() { ","versup() { alert('No se encontraron entradas del GLOSARIO en la pagina actual'); " )

	End If


	VerGlosario2 = XEhtml

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCION CONFECCION DEL EMAIL ( HTML O TEXTO)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GenMDLMAIL (XGUPARA,XGUORIGEN,XGUMENSAGE)

	''message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"


	XGUMENSAGEw = XGUMENSAGE & " "

	XGUMENSAGEw = Replace(XGUMENSAGEw,"@@@NORMREQ@@@","<u>" & SESSION("NOMREQ") & "</u>")
	XGUMENSAGEw = Replace(XGUMENSAGEw,"@@@FECREQ@@@",Now())



	Whtml = ""

	Select Case  SESSION("MAILOPC")

		Case "HTML"

		Whtml = Whtml & "<DIV STYLE='background-color:" & Session("CMFMENU") & ";height:40;width:600px;"
		Whtml = Whtml & "BORDER-color:#E0E0E0;BORDER-STYLE:SOLID;BORDER-WIDTH:1;'  >"
		' Whtml = Whtml & "<br>"
		Whtml = Whtml & "<P VALIGN=MIDDLE >&nbsp;&nbsp;<font size=4 face='" & Session("CMFMEFTFC") & "' color='" & Session("CMFMEFT") & "' ><B>"
		Whtml = Whtml & Session("CorTitApli") & "</B></FONT></P>"
		Whtml = Whtml & "</DIV>"
		Whtml = Whtml & "<BR>"

		Whtml = Whtml & "<DIV STYLE='BORDER-color:" & Session("CMFMENU") & ";BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
		Whtml = Whtml & "height:50;width:600px;'  >"
		Whtml = Whtml & "<CENTER>"

		Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
		Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
		Whtml = Whtml & "<P ALIGN=LEFT >"
		Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
		Whtml = Whtml & "&nbsp;&nbsp;DESTINATARIO</B></FONT></P>"
		Whtml = Whtml & "</DIV>"

		Whtml = Whtml & "<br>"
		Whtml = Whtml & "<DIV STYLE='width:80%;' >"
		Whtml = Whtml & "<P WIDTH=450 ALIGN='JUSTIFY' >"
		Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B><U>"
		Whtml = Whtml & XGUPARA & "</U></B></FONT></P>"
		Whtml = Whtml & "</DIV>"


		If request.querystring("OPWK") = "REASIGNAR" Then

			Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
			Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
			Whtml = Whtml & "<P ALIGN=LEFT >"
			Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
			Whtml = Whtml & "&nbsp;&nbsp;REASIGNADOR</B></FONT></P>"
			Whtml = Whtml & "</DIV>"

			xxxx     = getInfoUser(session("usuario"))
			xusrlog  = session("SGU_nombre")

			Whtml = Whtml & "<br>"
			Whtml = Whtml & "<DIV STYLE='width:80%;' >"
			Whtml = Whtml & "<P WIDTH=450 ALIGN='JUSTIFY' >"
			Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B><U>"
			Whtml = Whtml & xusrlog & "</U></B></FONT></P>"
			Whtml = Whtml & "</DIV>"

		End If

		If SOPCION = "RCH" Then

			Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
			Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
			Whtml = Whtml & "<P ALIGN=LEFT >"
			Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
			Whtml = Whtml & "&nbsp;&nbsp;RECHAZADOR</B></FONT></P>"
			Whtml = Whtml & "</DIV>"

			xxxx     = getInfoUser(session("usuario"))
			xusrlog  = session("SGU_nombre")

			Whtml = Whtml & "<br>"
			Whtml = Whtml & "<DIV STYLE='width:80%;' >"
			Whtml = Whtml & "<P WIDTH=450 ALIGN='JUSTIFY' >"
			Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B><U>"
			Whtml = Whtml & xusrlog & "</U></B></FONT></P>"
			Whtml = Whtml & "</DIV>"

		End If


		Whtml = Whtml & "<br>"
		Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
		Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
		Whtml = Whtml & "<P ALIGN=LEFT >"
		Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
		Whtml = Whtml & "&nbsp;&nbsp;DESCRIPCION</B></FONT></P>"
		Whtml = Whtml & "</DIV>"

		Whtml = Whtml & "<br>"
		Whtml = Whtml & "<DIV STYLE='width:80%;' >"
		Whtml = Whtml & "<P ALIGN='JUSTIFY' >"
		Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' >"
		Whtml = Whtml & XGUMENSAGEw
		Whtml = Whtml & "</FONT></P>"
		Whtml = Whtml & "</DIV>"
		Whtml = Whtml & "<BR>"
		Whtml = Whtml & "</DIV>"

		Whtml = Whtml & "<br>"
		Whtml = Whtml & "<DIV ALIGN='RIGHT' STYLE='width:200px;' >"
		Whtml = Whtml & "<P  >"
		Whtml = Whtml & "<font size=1 color=#6699cc face='" & Session("CMFMEFTFC") & "' ><B><i>"
		Whtml = Whtml & "COR&nbsp;E-SOLUTION&nbsp;SUITE&nbsp;V4.8.5</i></B></FONT></P>"
		Whtml = Whtml & "</DIV>"



		Case Else


		Whtml = Whtml & SESSION("NOMAPLI")  & "<BR>" & Chr(13) & Chr(13)

		Whtml = Whtml & XGUPARA & ": <BR>" & Chr(13)

		Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & Chr(13)

		Whtml = Whtml & XGUMENSAGEw  & "<BR>" & Chr(13)
		Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & Chr(13)




	End Select


	GenMDLMAIL = Whtml
	'RESPONSE.WRITE WHTML


End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCIONES DE BUSQUEDAS MULTIPLES ( USUARIOS-GRUPOS-SECTORES-EMPRESAS)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function  GenBuscaUsrex(XGUM1,XGUM2,XGUM3,XGUM4,XGUM5,XGUM6,XGUM7,XGUM8,XGUM9,XGUM10,XGUM1F,XGUM2F,XGUM3F,XGUM4F,XGUM5F,XGUM6F,XGUM7F,XGUM8F,XGUM9F,XGUM10F)
	'response.write request.querystring & "<br>" & request.form
	XGUM1FL = Replace(XGUM1F,Chr(9),"")
	XGUM2FL = Replace(XGUM2F,Chr(9),"")
	XGUM3FL = Replace(XGUM3F,Chr(9),"")
	XGUM4FL = Replace(XGUM4F,Chr(9),"")

	XGUM5FL = Replace(XGUM5F,Chr(9),"")
	XGUM6FL = Replace(XGUM6F,Chr(9),"")
	XGUM7FL = Replace(XGUM7F,Chr(9),"")
	XGUM8FL = Replace(XGUM8F,Chr(9),"")

	XGUM9FL = Replace(XGUM9F,Chr(9),"")
	XGUM10FL = Replace(XGUM10F,Chr(9),"")



	XGUFBURL = request.servervariables("SCRIPT_NAME")

	XGUFBURL = Right(XGUFBURL,Len(XGUFBURL) - (InStrRev(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

	XGUMASTER = REQUEST.FORM("XGUMASTER")
	XGUFILTRO = REQUEST.FORM("XGUFILTRO")
	XGUITEMCD = REQUEST.FORM("XGUITEMCD")
	XGUCODIGO = REQUEST.FORM("XGUCODIGO")
	XGUCODDSC = REQUEST.FORM("XGUCODDSC")
	'RESPONSE.WRITE REQUEST.FORM
	If XGUITEMCD="" Then XGUITEMCD="NO"
	If XGUCODIGO="" Then XGUCODIGO="NO"
	XGUMASTERW = XGUMASTER

	XGUMSBOT1  = "&nbsp;"
	XGUMSBOT2  = "&nbsp;"
	XGUMSBOT3  = "&nbsp;"
	XGUMSBOT4  = "&nbsp;"
	XGUMSBOT5  = "&nbsp;"
	XGUMSBOT6  = "&nbsp;"
	XGUMSBOT7  = "&nbsp;"
	XGUMSBOT8  = "&nbsp;"
	XGUMSBOT9  = "&nbsp;"
	XGUMSBOT10  = "&nbsp;"


	If SESSION("EMPMUN") = "S" Then

		If XGUM1 = "S" Then

			If  XGUMASTER = "" Or XGUMASTER = "NO" Then
				XGUMASTERW = "E"
			End If

			XGUMSBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESAS'  CLASS=BT ONCLICK=""clactfrm('E','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
			XGUMSBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI1 & "</FORM>"

		End If

	End If

	If XGUM2 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "S"
		End If

		XGUMSBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTORES'  CLASS=BT ONCLICK=""clactfrm('S','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI2 & "</FORM>"


	End If

	If XGUM3 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "G"
		End If

		XGUMSBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPOS'  CLASS=BT ONCLICK=""clactfrm('G','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI3 & "</FORM>"


	End If

	If XGUM4 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "U"
		End If

		XGUMSBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('U','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI4 & "</FORM>"


	End If

	If XGUM5 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "DEPENDENCIA"
		End If

		XGUMSBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('DEPENDENCIA','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI5 & "</FORM>"


	End If


	If XGUM6 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "PORPUESTO"
		End If

		XGUMSBOTi6 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORPUESTO','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT6  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI6 & "</FORM>"


	End If


	If XGUM7 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "PORUSR"
		End If

		XGUMSBOTi6 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORUSR','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT6  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI6 & "</FORM>"


	End If

	If XGUM8 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "PORSECPERF"
		End If

		XGUMSBOTi7 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORSECPERF','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT7  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI7 & "</FORM>"


	End If


	If XGUM9 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "PORSECTOREX"
		End If

		XGUMSBOTi8 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORSECTOREX','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT8  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI8 & "</FORM>"


	End If
	If XGUM10 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "PORGRUPOEXT"
		End If

		XGUMSBOTi9 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPOS DE CORREO'  CLASS=BT ONCLICK=""clactfrm('PORGRUPOEXT','NO','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>"
		XGUMSBOT9  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI9 & "</FORM>"


	End If






	If  XGUMASTER = "" Or XGUMASTER = "NO" Then
		XGUMASTER = XGUMASTERW
	End If


	XGUCONTEMS = ""

	XGUCONTEMS = XGUCONTEMS & "<table id=""MTBSUSR0B""  class='aaa'  cellspacing=1 cellPADDING=0 > "
	XGUCONTEMS = XGUCONTEMS & "<tr>"
	XGUCONTEMS = XGUCONTEMS & "<td class=ut WIDTH=60%  >&nbsp;</td>"

	If XGUMSBOT1 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT1
	End If

	If XGUMSBOT2 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT2
	End If

	If XGUMSBOT3 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT3
	End If

	If XGUMSBOT4 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT4
	End If


	If XGUMSBOT5 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT5
	End If

	If XGUMSBOT6 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT6
	End If

	If XGUMSBOT7 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT7
	End If

	If XGUMSBOT8 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT8
	End If
	If XGUMSBOT9 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT9
	End If


	XGUCONTEMS = XGUCONTEMS & "</tr>"
	XGUCONTEMS = XGUCONTEMS & "</table>"


	XGUFNC1 = ""
	XGUFNC2 = ""


	Set xgucn = ofv.conectar(ofv.strconn2)

	XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#F7AA22;cursor:pointer;' "
	XGUSTYLN = " STYLE='cursor:pointer;' "

	XGUFBUSR = XGUCODIGO

	XGUFBTAB1  = ""
	XGUFBTAB2  = ""
	XGUFBTD1  = ""
	XGUFBTD2  = ""

	If XGUMASTER = "" Or XGUMASTER = "NO" Then XGUMASTER = "U"

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' SELECCION DE BUSQUEDA MASTER - E:EMPRESAS / S:SECTORES / G:GRUPOS / U:USUARIOS              ''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' RESPONSE.WRITE XGUMASTER
	SESSION ("XGUMASTER" )=XGUMASTER

	Select Case XGUMASTER

		Case "PORUSR"


		XGUFILDSC = "&nbsp;POR&nbsp;NOMBRE"
		If  session("APLMDCLI") ="GP" Then
			XGUFBBOTi1 =""
			XGUFBBOTi2 =""
		Else
			XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
			XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		End If

		XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		If  session("APLMDCLI") ="GP" Then
			XGUFBBOT1  =""
			XGUFBBOT2  =""
		Else
			XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
			XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"

		End If

		XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
		XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"


		If XGUFILTRO="" Then XGUFILTRO=request.querystring("XGUFILTRO")
		If XGUITEMCD="" Then XGUITEMCD=request.querystring("XGUITEMCD")


		Select Case XGUFILTRO




			Case "L"

			XGUFILDSC1 = ":&nbsp;LETRA"

			xTrColor = "#F9F9F9"


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If


			Dim ATABXGULTRz(28)

			ATABXGULTRz(1)  = "A"
			ATABXGULTRz(2)  = "B"
			ATABXGULTRz(3)  = "C"
			ATABXGULTRz(4)  = "D"
			ATABXGULTRz(5)  = "E"
			ATABXGULTRz(6)  = "F"
			ATABXGULTRz(7)  = "G"
			ATABXGULTRz(8)  = "H"
			ATABXGULTRz(9)  = "I"
			ATABXGULTRz(10) = "J"
			ATABXGULTRz(11) = "K"
			ATABXGULTRz(12) = "L"
			ATABXGULTRz(13) = "M"
			ATABXGULTRz(14) = "N"
			ATABXGULTRz(15) = "O"
			ATABXGULTRz(16) = "P"
			ATABXGULTRz(17) = "Q"
			ATABXGULTRz(18) = "R"
			ATABXGULTRz(19) = "S"
			ATABXGULTRz(20) = "T"
			ATABXGULTRz(21) = "U"
			ATABXGULTRz(22) = "V"
			ATABXGULTRz(23) = "W"
			ATABXGULTRz(24) = "X"
			ATABXGULTRz(25) = "Y"
			ATABXGULTRz(26) = "Z"
			ATABXGULTRz(27) = "NO"
			ATABXGULTRz(28) = "NO"


			For XGUINX=1 To UBound(ATABXGULTRz) Step 4

				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If ATABXGULTRz(XGUINX) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & "</TD>"


				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If ATABXGULTRz(XGUINX+1) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " >"
				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & "</TD>"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If  ATABXGULTRz(XGUINX+2) = "NO" Then

					XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

				Else
					If ATABXGULTRz(XGUINX+2) = XGUITEMCD Then
						XGUFBTD1 = XGUFBTD1 & XGUSTYLS
					Else
						XGUFBTD1 = XGUFBTD1 & XGUSTYLN
					End If
					XGUFBTD1 = XGUFBTD1 & " >"
					XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
					XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

				End If

				XGUFBTD1 = XGUFBTD1 & "</TD>"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If  ATABXGULTRz(XGUINX+3) = "NO" Then

					XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

				Else
					If ATABXGULTRz(XGUINX+3) = XGUITEMCD Then
						XGUFBTD1 = XGUFBTD1 & XGUSTYLS
					Else
						XGUFBTD1 = XGUFBTD1 & XGUSTYLN
					End If
					XGUFBTD1 = XGUFBTD1 & " >"
					XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
					XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

				End If

				XGUFBTD1 = XGUFBTD1 & "</TD>"



				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If

			Next




			If XGUITEMCD <> "NO" Then

				Call getUsrInTEXT2(XGUITEMCD & "*")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If




			Case "T"
			' XGUITEMCD=replace(XGUITEMCD &"","'","%27")
			XGUFILDSC1 = ":&nbsp;TEXTO"



			xTrColor = "#FFFFFF"


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If



			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

			If XGUITEMCD <> "NO" Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"
			XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE=""" & XGUITEMCD& """ CLASS=DT size=40 >&nbsp;&nbsp;"

			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If XGUITEMCD <> "NO" Then

				Call getUsrInTEXT2("*" & XGUITEMCD & "*")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & Replace(XGUITEMCD &"","'","\'")  & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If

			Case Else

			XGUFILDSC1 = "&nbsp;"

			XGUFILTRO = "NO"
			XGUITEMCD = "NO"
			XGUCODIGO = "NO"
			XGUCODDSC = "NO"

		End Select

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;USUARIOS</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"

		Case "DEPENDENCIA"

		XGUFILDSC = "&nbsp;POR&nbsp;DEPENDENCIA"

		xTrColor = "#FFFFFF"
		'RESPONSE.WRITE XGUITEMCD & "22"
		If XGUITEMCD <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		strsql = "Select e.* From Eco e "
		strsql =strsql & "where e.eco in(select distinct l.sector from listadepuestos l where l.grafica='S' and l.site='" & SESSION("SUN") & "')"
		strsql =strsql & " order by e.descri"
		''RESPONSE.WRITE strsql & " strsql "
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

			If CStr(xgurs(0)) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)

		session("SGU_UIGuserid")=""
		session("SGU_UIGdescri")=""
		If XGUITEMCD <> "NO" Then

			' CALL getUsrInSECT(XGUITEMCD,"")
			If XGUITEMCD=0 Then XGUITEMCD=-1
			SQL="SELECT * FROM LISTADEPUESTOS WHERE site='" & SESSION("SUN") & "' and grafica='S' and sector=" &  XGUITEMCD
			'response.write sql
			Set xpuestos = ofv.crearconsultaEx(Sql,xgucn,1,volver)
			delim=""

			Do While Not xpuestos.eof

				session("SGU_UIGuserid")= session("SGU_UIGuserid")   & delim  & xpuestos(0)
				session("SGU_UIGdescri")= session("SGU_UIGdescri")  & delim  & xpuestos(1)
				delim =Chr(9)
				xpuestos.movenext
			Loop
			Call  ofv.cerrarconsulta(xpuestos)
			XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
			XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

			xTrColor = "#FFFFFF"

			For XGUINX=0 To UBound(XGUTABKEY)


				XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD2 = XGUFBTD2 & " >"

				XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR PUESTO.' ID='---" & XGUTABKEY(XGUINX) & "' "

				If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
					XGUFBTD2 = XGUFBTD2 & XGUSTYLS
				Else
					XGUFBTD2 = XGUFBTD2 & XGUSTYLN
				End If

				XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
				XGUFBTD2 = XGUFBTD2 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If


			Next




		End If
		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;PUESTO</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


		Case "PORSECTOREX"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""

		XGUFILDSC = "&nbsp;SECTOR"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If



		STRSQL= "Select A.OBJID,B.DESCRI From objsector A"
		STRSQL = STRSQL & " INNER JOIN ECO B ON(B.ECO= A.CODSECTOR)"
		STRSQL=STRSQL & " WHERE A.CODOBJETO=0" & session("NCliente")
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR SECTOR.'  ID='***" & xgurs(0) & "' "

			If (XGUM2FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM2FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		Case "PORGRUPOEXT"




		XGUFILDSC = "&nbsp;POR&nbsp;GRUPOS DE CORREO"
		If  session("APLMDCLI") ="GP" Then
			XGUFBBOTi1 =""' "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
			XGUFBBOTi2 =""' "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		Else
			XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
			XGUFBBOTi2 ="<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		End If

		' XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		If  session("APLMDCLI") ="GP" Then
			XGUFBBOT1  =""' "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
			XGUFBBOT2  =""' "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
		Else
			XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
			XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"

		End If

		' XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
		XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"


		If XGUFILTRO="" Then XGUFILTRO=request.querystring("XGUFILTRO")
		If XGUITEMCD="" Then XGUITEMCD=request.querystring("XGUITEMCD")


		Select Case XGUFILTRO


			Case "T"
			' XGUITEMCD=replace(XGUITEMCD &"","'","%27")
			XGUFILDSC1 = ":&nbsp;TEXTO"



			xTrColor = "#FFFFFF"


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If



			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECIONAR FILTRO.'  "

			If XGUITEMCD <> "NO" Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"
			XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE=""" & XGUITEMCD& """ CLASS=DT size=40 >&nbsp;&nbsp;"

			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If XGUITEMCD <> "NO" Then

				Call getGruposMail("*" & XGUITEMCD & "*")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR GRUPO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & Replace(XGUITEMCD &"","'","\'")  & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If

			Case Else

			XGUFILDSC1 = "&nbsp;"

			XGUFILTRO = "NO"
			XGUITEMCD = "NO"
			XGUCODIGO = "NO"
			XGUCODDSC = "NO"

		End Select

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;GRUPO</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


		Case "PORSECPERF"

		XGUFILDSC = "&nbsp;SECTORES/PERFIL"

		xTrColor = "#FFFFFF"
		'RESPONSE.WRITE XGUITEMCD & "22"
		If XGUITEMCD <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		' strsql = "Select e.* From Eco e "
		' strsql =strsql & "where e.eco in(select distinct l.codsector from objsector l where l.codobjeto=0"& session("NCliente") &")"
		' strsql =strsql & " order by e.descri"
		strsql= " Select A.OBJID,B.DESCRI From objsector A "
		strsql=strsql & " INNER JOIN ECO B ON(B.ECO= A.CODSECTOR)"
		strsql=strsql & " WHERE A.CODOBJETO=0" & session("NCliente")
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

			If CStr(xgurs(0)) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)
		' ver ale
		session("SGU_UIGuserid")=""
		session("SGU_UIGdescri")=""
		If XGUITEMCD <> "NO" Then

			' CALL getUsrInSECT(XGUITEMCD,"")
			If XGUITEMCD=0 Then XGUITEMCD=-1
			'sql = "Select Ecoperfil.*,Tipoperfil.descri From Ecoperfil, Tipoperfil where "
			'sql = sql & "Ecoperfil.codperfil = Tipoperfil.tipoperfil and codsector=" & XGUITEMCD

			SQL= "Select A.*,B.descri From Ecoperfil A"
			SQL=SQL & " INNER JOIN Tipoperfil B ON( A.codperfil=B.tipoperfil)"
			SQL=SQL & " INNER JOIN  objsector C ON (C.OBJID=0"&XGUITEMCD&" AND C.CODSECTOR=A.CODSECTOR)"

			Set xpuestos = ofv.crearconsultaEx(Sql,xgucn,1,volver)
			delim=""

			Do While Not xpuestos.eof

				session("SGU_UIGuserid")= session("SGU_UIGuserid")   & delim  & xpuestos(0)
				session("SGU_UIGdescri")= session("SGU_UIGdescri")  & delim  & xpuestos(4)
				delim =Chr(9)
				xpuestos.movenext
			Loop
			Call  ofv.cerrarconsulta(xpuestos)
			XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
			XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

			xTrColor = "#FFFFFF"

			For XGUINX=0 To UBound(XGUTABKEY)


				XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD2 = XGUFBTD2 & " >"

				XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR PUESTO.' ID='---" & XGUTABKEY(XGUINX) & "' "

				If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
					XGUFBTD2 = XGUFBTD2 & XGUSTYLS
				Else
					XGUFBTD2 = XGUFBTD2 & XGUSTYLN
				End If

				XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
				XGUFBTD2 = XGUFBTD2 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If


			Next




		End If
		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:150;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:150;width:99%;background-color:#ffffff;"
		XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;PUESTO</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"



		Case "PORPUESTO"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""


		XGUFILDSC = "&nbsp;POR&nbsp;PUESTO"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If

		strsql = "SELECT B.OBJID,B.DESCRI + ' - ' + C.DESCRI collate database_default AS T1 FROM LISTADEPUESTOS B  INNER JOIN ECO C ON (b.site='" & SESSION("SUN") &"' and C.ECO=B.SECTOR )  ORDER BY T1  "
		' response.write strsql
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR PUESTO.'  ID='***" & xgurs(0) & "' "

			If (XGUM3FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM3FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"




		Case "E"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""

		XGUFILDSC = "&nbsp;EMPRESA"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If session("UN") = 1 Then
			strsql = "Select * From clienteplan  order by cliente"
		Else
			' 02/09

			XEMPRE = Split(session("EmpHbl"),Chr(9))

			DELIM=""
			idempre = ""

			For i=0 To UBound(XEMPRE)

				idempre = idempre & DELIM & XEMPRE(i)
				DELIM = ","

			Next
			strsql = "Select * From clienteplan where id in (" & idempre & ")  order by cliente"
		End If

		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR EMPRESA.' ID='***" & xgurs(0) & "' "

			If (XGUM1FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM1FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


		'''' Whtml = Whtml & " document.getElementById(spmk).scrollIntoView(true); " & chr(13)




		Case "S"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""

		XGUFILDSC = "&nbsp;SECTOR"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If



		strsql    = "Select * From Eco  order by descri"
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR SECTOR.'  ID='***" & xgurs(0) & "' "

			If (XGUM2FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM2FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


		Case "G"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""


		XGUFILDSC = "&nbsp;GRUPO"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If

		' strsql = "Select * From GRUPOS  order by descriPCION "
		If session("USRPRV")  = "I" Or session("USRPRV")  = "E" Then
			strsql = "Select * From GRUPOS  order by descriPCION "
		Else
			strsql = "Select * From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
		End If


		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR GRUPO.'  ID='***" & xgurs(0) & "' "

			If (XGUM3FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM3FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"






		Case "U"




		'XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>"
		XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>"
		XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>"
		XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>"

		If XGUM2 <> "N" Then
			XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		Else
			XGUFBBOTi2 = "<td class=ut  >&nbsp;</td>"
		End If
		'	XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		'	XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

		XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
		XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
		XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
		XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"

		If REQUEST.FORM("XGUITEMCD") =  ""  Or  Not IsNumeric(REQUEST.FORM("XGUITEMCD")) Then
			If session("SETIDEMPRE") = "1" Or session("SETIDEMPRE") = "" Then session("SETIDEMPRE")=  SESSION("CLIEMP")
		Else
			session("SETIDEMPRE")=  REQUEST.FORM("XGUITEMCD")
		End If


		strsql = "Select CLIENTE From CLIENTEPLAN WHERE ID=" & session("SETIDEMPRE")
		'response.write strsql
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		If Not xgurs.EOF Then
			Scliente  = " : " & xgurs(0)
		End If
		Call OFV.CERRARCONSULTA(xgurs)

		If Session("EMPMUN")  = "S" Then

			strsql = "Select *  From DOMINIOS WHERE cliente =" & session("SETIDEMPRE")
			'response.write  strsql
			Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

			If Not xgurs.EOF Then
				session("ldapEmpre")=xgurs("raizldap")
				session("usrEmpre")=xgurs("usrldap")
				session("pwdEmpre")=xgurs("pwdldap")
			End If
			Call OFV.CERRARCONSULTA(xgurs)
		End If

		XGUFILDSC = "&nbsp;POR&nbsp;USUARIO"

		Select Case XGUFILTRO

			Case "G"

			XGUFILDSC1 = ":&nbsp;GRUPO"

			xTrColor = "#FFFFFF"

			If XGUITEMCD <> "NO" Then
				XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC1 = ""
			End If


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If


			If session("USRPRV")  = "I"  Then
				strsql = "Select ID,descripcion From GRUPOS  order by descriPCION "  ' ok
				'strsql = "Select mapgrp,ID From GRUPOS  order by descriPCION " ' probar

			Else
				strsql = "Select ID,descripcion From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
			End If

			Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)
			delim=""
			sgrupos=""


			Do While Not xgurs.eof
				sgrupos = sgrupos & delim & RTrim(LTrim(xgurs(0) & "" ))
				delim = Chr(9)
				xgurs.movenext
			Loop
			Call ofv.cerrarconsulta(xgurs.eof)
			xx= buscarGrupoAd(sgrupos)


			XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
			XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

			xTrColor = "#FFFFFF"

			For XGUINX=0 To UBound(XGUTABKEY)



				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.' ID='***" & XGUTABKEY(XGUINX) & "' "

				If CStr(XGUTABKEY(XGUINX)) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If



				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUTABKEY(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & XGUTABDAT(XGUINX) & "</TD>"
				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If
			Next

			If XGUITEMCD <> "NO" Then

				Call getUsrIngroup(XGUITEMCD)

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				WWWW = InStr(UCase(REQUEST.SERVERVARIABLES("SCRIPT_NAME") & ""),"AGREGARDIST.ASP")

				If WWWW > 0 Then

					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor='#6699CC' "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=VTI STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#000000;COLOR:#FFFFFF;' "
					XGUFBTD2 = XGUFBTD2 & " TITLE='SELECCIONAR TODOS LOS USUARIOS.'  ID='---XXXTODOSXXX" & XGUITEMCD & "' "

					If (XGUM4FL & "TODOS") = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & "XXXTODOSXXX" & XGUITEMCD & "','" & "TODOS LOS USUARIOS" & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & "TODOS LOS USUARIOS" & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"

				End If


				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If



			Case "S"

			XGUFILDSC1 = ":&nbsp;SECTOR"

			xTrColor = "#FFFFFF"

			If XGUITEMCD <> "NO" Then
				XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC1 = ""
			End If


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If



			strsql = "Select * From Eco  order by descri"
			Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

			Do Until xgurs.eof


				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

				If CStr(xgurs(0)) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If


				xgurs.movenext
			Loop

			Call ofv.cerrarconsulta(xgurs)


			If XGUITEMCD <> "NO" Then

				Call getUsrInSECT(XGUITEMCD,"")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If




			Case "L"

			XGUFILDSC1 = ":&nbsp;LETRA"

			xTrColor = "#F9F9F9"


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If


			Dim ATABXGULTR(28)

			ATABXGULTR(1)  = "A"
			ATABXGULTR(2)  = "B"
			ATABXGULTR(3)  = "C"
			ATABXGULTR(4)  = "D"
			ATABXGULTR(5)  = "E"
			ATABXGULTR(6)  = "F"
			ATABXGULTR(7)  = "G"
			ATABXGULTR(8)  = "H"
			ATABXGULTR(9)  = "I"
			ATABXGULTR(10) = "J"
			ATABXGULTR(11) = "K"
			ATABXGULTR(12) = "L"
			ATABXGULTR(13) = "M"
			ATABXGULTR(14) = "N"
			ATABXGULTR(15) = "O"
			ATABXGULTR(16) = "P"
			ATABXGULTR(17) = "Q"
			ATABXGULTR(18) = "R"
			ATABXGULTR(19) = "S"
			ATABXGULTR(20) = "T"
			ATABXGULTR(21) = "U"
			ATABXGULTR(22) = "V"
			ATABXGULTR(23) = "W"
			ATABXGULTR(24) = "X"
			ATABXGULTR(25) = "Y"
			ATABXGULTR(26) = "Z"
			ATABXGULTR(27) = "NO"
			ATABXGULTR(28) = "NO"


			For XGUINX=1 To UBound(ATABXGULTR) Step 4

				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If ATABXGULTR(XGUINX) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & "</TD>"


				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If ATABXGULTR(XGUINX+1) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " >"
				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & "</TD>"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If  ATABXGULTR(XGUINX+2) = "NO" Then

					XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

				Else
					If ATABXGULTR(XGUINX+2) = XGUITEMCD Then
						XGUFBTD1 = XGUFBTD1 & XGUSTYLS
					Else
						XGUFBTD1 = XGUFBTD1 & XGUSTYLN
					End If
					XGUFBTD1 = XGUFBTD1 & " >"
					XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
					XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

				End If

				XGUFBTD1 = XGUFBTD1 & "</TD>"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If  ATABXGULTR(XGUINX+3) = "NO" Then

					XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

				Else
					If ATABXGULTR(XGUINX+3) = XGUITEMCD Then
						XGUFBTD1 = XGUFBTD1 & XGUSTYLS
					Else
						XGUFBTD1 = XGUFBTD1 & XGUSTYLN
					End If
					XGUFBTD1 = XGUFBTD1 & " >"
					XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
					XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

				End If

				XGUFBTD1 = XGUFBTD1 & "</TD>"



				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If

			Next




			If XGUITEMCD <> "NO" Then

				Call getUsrInTEXT2(XGUITEMCD & "*")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If




			Case "T"

			XGUFILDSC1 = ":&nbsp;TEXTO"



			xTrColor = "#FFFFFF"


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If



			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

			If XGUITEMCD <> "NO" Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"
			XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE=""" & XGUITEMCD & """ CLASS=DT size=40 >&nbsp;&nbsp"

			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If XGUITEMCD <> "NO" Then

				Call getUsrInTEXT2("*" & XGUITEMCD & "*")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" &  Replace(XGUITEMCD &"","'","\'") & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If






			Case Else

			XGUFILDSC1 = "&nbsp;"

			XGUFILTRO = "NO"
			XGUITEMCD = "NO"
			XGUCODIGO = "NO"
			XGUCODDSC = "NO"


		End Select






		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;USUARIOS</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


	End Select


	Call ofv.cerrarconn(xgucn)


	Whtml = ""

	Whtml = Whtml & "<center>"

	If XGUMASTER="DEPENDENCIA" Or XGUMASTER="PORPUESTO" Then
		XANCHO="95%"
	Else
		XANCHO="60%"
	End If
	Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:"&XANCHO&";"
	Whtml = Whtml & "'>"
	Whtml = Whtml & "<CENTER>"

	Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=99% cellspacing=0 cellpadding=2 >"
	Whtml = Whtml & "<tr HEIGHT=25 >"
	Whtml = Whtml & "<td class=ttr width=99%  >&nbsp;&nbsp;BUSCAR" & XGUFILDSC & "&nbsp;</td>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR0""  class='aaa' width=99%  cellspacing=0 cellpadding=0 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"
	Whtml = Whtml & "<tD ALIGN=RIGHT >"
	Whtml = Whtml & XGUCONTEMS
	Whtml = Whtml & "</tD>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR2""  class='aaa' width=99%  cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"
	Whtml = Whtml & XGUFBBOT1
	Whtml = Whtml & XGUFBBOT2
	Whtml = Whtml & XGUFBBOT3
	Whtml = Whtml & XGUFBBOT4
	Whtml = Whtml & XGUFBBOT5

	If XGUCODDSC <> "NO" Then
		Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
	Else
		Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
	End If

	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR3""   class='aaa' width=99% cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"

	Whtml = Whtml & XGUCONTENT

	Whtml = Whtml & "</table>"
	Whtml = Whtml & "</CENTER>"
	Whtml = Whtml & "</DIV>"
	Whtml = Whtml & "</center>"

	If XGUMASTERW="PORSECPERF" Or  XGUMASTERW="PORSECTOREX" Then
		Whtml = Whtml & "<FORM NAME='XGUFRM' id='XGUFRM' TARGET=WK4 ACTION='" & XGUFBURL & "' METHOD='POST' >"
	Else
		Whtml = Whtml & "<FORM NAME='XGUFRM'  id='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' >"
	End If
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' id='XGUMASTER' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFILTRO' id='XGUFILTRO'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUITEMCD' id='XGUITEMCD'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODIGO' id='XGUCODIGO'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODDSC' id='XGUCODDSC'>"
	Whtml = Whtml & "</FORM>"

	Whtml = Whtml & "<script language=javascript >"
	Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4) { " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU1; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU2;" & Chr(13)
	'Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=document.getElementById('XGUITEMCD').value.replace('\'','fxx');" & chr(13)
	Whtml = Whtml & "    document.getElementById('XGUCODIGO').value=XGU3; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUCODDSC').value=XGU4; " & Chr(13)
	Whtml = Whtml & " } " & Chr(13)
	Whtml = Whtml & XGUFNC1
	Whtml = Whtml & XGUFNC2

	Whtml = Whtml & "</script>"



	SESSION("FBUSR") = XGUCODIGO
	SESSION("FBSLC") = XGUMASTER
	session("XGUCODDSC")=XGUCODDSC
	SESSION("XGUITEMCD")=XGUITEMCD
	SESSION("XGUFILTRO")=XGUFILTRO

	GenBuscaUsrEx = Whtml




End Function


Function GenBuscaUsr(XGUM1,XGUM2,XGUM3,XGUM4,XGUM1F,XGUM2F,XGUM3F,XGUM4F)
	'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form & "<br>"

	XGUM1FL = Replace(XGUM1F,Chr(9),"")
	XGUM2FL = Replace(XGUM2F,Chr(9),"")
	XGUM3FL = Replace(XGUM3F,Chr(9),"")
	XGUM4FL = Replace(XGUM4F,Chr(9),"")


	XGUFBURL = request.servervariables("SCRIPT_NAME")

	XGUFBURL = Right(XGUFBURL,Len(XGUFBURL) - (InStrRev(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

	XGUMASTER = REQUEST.FORM("XGUMASTER")
	XGUFILTRO = REQUEST.FORM("XGUFILTRO")
	XGUITEMCD = REQUEST.FORM("XGUITEMCD")
	XGUCODIGO = REQUEST.FORM("XGUCODIGO")
	XGUCODDSC = REQUEST.FORM("XGUCODDSC")


	XGUMASTERW = XGUMASTER

	XGUMSBOT1  = "&nbsp;"
	XGUMSBOT2  = "&nbsp;"
	XGUMSBOT3  = "&nbsp;"
	XGUMSBOT4  = "&nbsp;"
	XGUMSBOT5  = "&nbsp;"


	'RESPONSE.WRITE REQUEST.FORM
	If SESSION("EMPMUN") = "S" Then

		If XGUM1 = "S" Then

			If  XGUMASTER = "" Or XGUMASTER = "NO" Then
				XGUMASTERW = "E"
			End If

			XGUMSBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESAS'  CLASS=BT ONCLICK=""clactfrm('E','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
			XGUMSBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI1 & "</FORM>"

		End If

	End If

	If XGUM2 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "S"
		End If

		'XGUMSBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTORES'  CLASS=BT ONCLICK=""clactfrm('S','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI2 & "</FORM>"


	End If

	If XGUM3 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "G"
		End If

		XGUMSBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPOS'  CLASS=BT ONCLICK=""clactfrm('G','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI3 & "</FORM>"


	End If



	If XGUM4 = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERW = "U"
		End If
		If request.querystring("OPWK") <> "REASIGNAR" Then
			XGUMSBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='USUARIOS'  CLASS=BT ONCLICK=""clactfrm('U','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
			XGUMSBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTi4 & "</FORM>"
		End If

	End If

	If SESSION("EMPMUN") = "S" Then

		If  XGUMASTER = "" Or XGUMASTER = "NO" Then
			XGUMASTERWT = "TE"
		End If

		XGUMSBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESAS'  CLASS=BT ONCLICK=""clactfrm('TE','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUMSBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTi5 & "</FORM>"


	End If

	If  XGUMASTER = "" Or XGUMASTER = "NO" Then
		XGUMASTER = XGUMASTERW
	End If


	XGUCONTEMS = ""

	XGUCONTEMS = XGUCONTEMS & "<table   id=""MTBSUSR0B""  class=''  cellspacing=1 cellPADDING=0 > "
	XGUCONTEMS = XGUCONTEMS & "<tr>"
	'XGUCONTEMS = XGUCONTEMS & "<td class=ut WIDTH=60%  >&nbsp;</td>"

	If XGUMSBOT1 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT1
	End If

	If XGUMSBOT2 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS  & XGUMSBOT2
	End If

	If XGUMSBOT3 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT3
	End If

	If XGUMSBOT4 <> "&nbsp;" Then
		XGUCONTEMS = XGUCONTEMS & XGUMSBOT4
	End If


	XGUCONTEMS = XGUCONTEMS & "</tr>"
	XGUCONTEMS = XGUCONTEMS & "</table>"


	XGUFNC1 = ""
	XGUFNC2 = ""


	Set xgucn = ofv.conectar(ofv.strconn2)

	XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";cursor:pointer;' "
	XGUSTYLN = " STYLE='cursor:pointer;' "

	XGUFBUSR = XGUCODIGO

	XGUFBTAB1  = ""
	XGUFBTAB2  = ""
	XGUFBTD1  = ""
	XGUFBTD2  = ""

	If XGUMASTER = "" Or XGUMASTER = "NO" Then XGUMASTER = "U"

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' SELECCION DE BUSQUEDA MASTER - E:EMPRESAS / S:SECTORES / G:GRUPOS / U:USUARIOS              ''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	SESSION ("XGUMASTER" )=XGUMASTER
	'	RESPONSE.WRITE XGUMASTER
	Select Case XGUMASTER



		Case "E"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""

		XGUFILDSC = "&nbsp;EMPRESA"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If

		XEMPRE = Split(session("EmpHbl"),Chr(9))

		DELIM=""
		idempre = ""

		For i=0 To UBound(XEMPRE)

			idempre = idempre & DELIM & XEMPRE(i)
			DELIM = ","

		Next

		If idempre = "" Then idempre = session("CLIEMP")

		strsql = "Select ID,CLIENTE From CLIENTEPLAN where id in (" & idempre & ") order by ID "
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR EMPRESA.' ID='***" & xgurs(0) & "' "

			If (XGUM1FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM1FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


		'''' Whtml = Whtml & " document.getElementById(spmk).scrollIntoView(true); " & chr(13)
		'ACA

		Case "S"

		XGUFILDSC1 = ":&nbsp;EMPRESAS"

		xTrColor = "#FFFFFF"

		If XGUITEMCD <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If


		XEMPRE = Split(session("EmpHbl"),Chr(9))

		DELIM=""
		idempre = ""

		For i=0 To UBound(XEMPRE)

			'idempre = idempre & DELIM & XEMPRE(i)
			DELIM = ","

		Next
		If idempre = "" Then idempre = session("UN") 'session("CLIEMP")

		'strsql = "Select * From clienteplan  where id in(" & idempre & ") "
		strsql ="Select * From clienteplan  where id in(select objun from objetos where CODOBJETO=0" &  session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "') "
		'where  CODOBJETO=0" &  session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "'
		'response.write strsql
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

			If CStr(xgurs(0)) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)

		XEMPRE = Split(session("EmpHbl"),Chr(9))

		DELIM=""
		idempre = ""

		For i=0 To UBound(XEMPRE)

			idempre = idempre & DELIM & XEMPRE(i)
			DELIM = ","

		Next

		If idempre = "" Then idempre = session("CLIEMP")

		If XGUITEMCD = 0 Then XGUITEMCD = session("UN")

		If XGUITEMCD <> "NO" Then

			'	Call getUsrInSECT(XGUITEMCD,"")


			'	If XGUITEMCD <> 1 Then
			sql="select  eco,descri from eco where estado=3 and ecoun=0" & XGUITEMCD & " order by descri"
			'	Else
			'		sql="select  eco,descri from eco where estado=3 and ecoun in (" & idempre & ") order by descri"
			'	End If
			'response.write sql
			Set xgurs = ofv.crearconsultaEx(sql,xgucn,1,volver)
			delim=""
			session("SGU_UIGuserid")=""
			session("SGU_UIGdescri")=""
			Do While Not xgurs.eof
				session("SGU_UIGuserid")=session("SGU_UIGuserid") & delim & xgurs(0)
				session("SGU_UIGdescri")=session("SGU_UIGdescri") & delim & xgurs(1)
				delim= Chr(9)
				xgurs.movenext
			Loop
			Call ofv.cerrarconsulta(xgurs)

			XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
			XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

			xTrColor = "#FFFFFF"

			For XGUINX=0 To UBound(XGUTABKEY)


				XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD2 = XGUFBTD2 & " >"

				XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR SECTOR.' ID='---" & XGUTABKEY(XGUINX) & "' "

				If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
					XGUFBTD2 = XGUFBTD2 & XGUSTYLS
				Else
					XGUFBTD2 = XGUFBTD2 & XGUSTYLN
				End If

				XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
				XGUFBTD2 = XGUFBTD2 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If


			Next




		End If
		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
		XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		If  Session("EMPMUN") = "S"  Then
			XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		Else
			XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% STYLE='DISPLAY:NONE' >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		End If
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SECTORES</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		If  Session("EMPMUN") = "S"  Then
			XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		Else
			XGUCONTENT = XGUCONTENT & "<td class=ut  STYLE='DISPLAY:NONE' >" & XGUFBDIV1 & "</td>"

		End If

		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


		Case "Se"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""

		XGUFILDSC = "&nbsp;SECTOR"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If



		strsql    = "Select * From Eco  order by descri"
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR SECTOR.'  ID='***" & xgurs(0) & "' "

			If (XGUM2FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM2FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


		Case "G"

		XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
		XGUFBBOT2  = ""
		XGUFBBOT3  = ""
		XGUFBBOT4  = ""


		XGUFILDSC = "&nbsp;GRUPO"

		xTrColor = "#FFFFFF"

		If XGUCODIGO <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If

		' strsql = "Select * From GRUPOS  order by descriPCION "
		If session("USRPRV")  = "I" Or session("USRPRV")  = "E" Then
			strsql = "Select * From GRUPOS  order by descriPCION "
		Else
			strsql = "Select * From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
		End If

		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECIONAR GRUPO.'  ID='***" & xgurs(0) & "' "

			If (XGUM3FL & CStr(xgurs(0))) = XGUCODIGO Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM3FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:98%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"






		Case "U"

		If request.querystring("OPWK") <> "REASIGNAR" Then
			If session("APLMDCLI")  = "GP" Then

				'XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
				'XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
				XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
				XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"



				If REQUEST.FORM("XGUITEMCD") =  ""  Or  Not IsNumeric(REQUEST.FORM("XGUITEMCD")) Then
					'if session("SETIDEMPRE") = "1" or session("SETIDEMPRE") = "" then session("SETIDEMPRE")=  SESSION("CLIEMP")
					If session("SETIDEMPRE") = "" Then session("SETIDEMPRE")=  SESSION("CLIEMP")
				Else
					session("SETIDEMPRE")=  REQUEST.FORM("XGUITEMCD")
				End If


				strsql = "Select CLIENTE From CLIENTEPLAN WHERE ID=" & session("SETIDEMPRE")
				Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

				If Not xgurs.EOF Then
					Scliente  = " : " & xgurs(0)
				End If
				Call OFV.CERRARCONSULTA(xgurs)

				If Session("EMPMUN")  = "S" Then

					strsql = "Select *  From DOMINIOS WHERE cliente =" & session("SETIDEMPRE")
					'response.write  strsql
					Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

					If Not xgurs.EOF Then
						session("ldapEmpre")=xgurs("raizldap")
						session("usrEmpre")=xgurs("usrldap")
						session("pwdEmpre")=xgurs("pwdldap")
					End If
					Call OFV.CERRARCONSULTA(xgurs)
				End If
				If Session("EMPMUN") = "S" Then
					XGUFBBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESA" & Scliente  &  "'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','TE','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
				Else
					XGUFBBOTi5=""
				End If



				XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
				XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
				XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
				XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"
				XGUFBBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI5 & "</FORM>"
			Else



				XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
				XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
				XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
				XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"



				XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
				XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
				XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
				XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"
			End If
		Else
			XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
			XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
			XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
			XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"
		End If
		XGUFILDSC = "&nbsp;USUARIO"

		Select Case XGUFILTRO

			Case "TEFND"

			' XGUFNC3 = "alert('1');clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"
			Case "TE"
			'response.write "empXGUITEMCD: " & XGUITEMCD & " empXGUCODIGO: " & XGUCODIGO
			XGUFILDSC1 = ":&nbsp;EMPRESA"

			xTrColor = "#FFFFFF"

			If XGUITEMCD <> "NO" Then
				session("ssec") = XGUITEMCD
				XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC1 = ""
			End If


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If

			' 02/08/2021 aca se agrega la nueva funcionalidad OBJMEMPALC -> ALE

			If session("UN") = 1 Then

				XEMPRE = SPLIT(session("EmpHbl"),chr(9))

				DELIM=""
				idempre = ""

				For i=0 To ubound(XEMPRE)

					idempre = idempre & DELIM & XEMPRE(i)
					DELIM = ","

				Next


			End If

			If idempre = "" Then idempre = session("CLIEMP")

			'       strsql = "Select ID,CLIENTE From CLIENTEPLAN where id in (" & idempre & ") order by ID "

			'else
			'response.write XGUFBURL
			If instr(ucase(XGUFBURL),ucase("AgregarCircuitoESQUEMA.asp")) = 0 Then
				sqlac="select distinct CODEMPRE from OBJMEMPALC where  CODOBJETO=0" &  session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
				sqlac = sqlac & " and CODEMPRE not in (" & idempre & ")"
				'   response.write XGUFBURL
				Set acRs = ofv.crearconsultaEx(sqlac,xgucn,1,volver)


				Do While Not acRs.eof
					DELIM = ","
					idempre = idempre & DELIM & acRs(0)

					acRs.movenext
				Loop
				Call ofv.cerrarconsulta(acRs)
			End If
			'strsql = "Select ID,CLIENTE From CLIENTEPLAN where id=" & session("UN")

			' 02/08/2021 aca se agrega la nueva funcionalidad OBJMEMPALC -> ALE
			'10317-0004/20
			strsql = "Select ID,CLIENTE From CLIENTEPLAN where id in (" & idempre & ") order by ID "


			'end if

			sqlac="select objun from objetos where  CODOBJETO=0" &  session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
			Set acRs = ofv.crearconsultaEx(sqlac,xgucn,1,volver)
			'revisar
			If Not  acRs.eof Then
				If acRs(0) <> 1 Then  strsql = "Select ID,CLIENTE From CLIENTEPLAN where id=" & session("UN")
			End If
			Call ofv.cerrarconsulta(acRs)

			' strsql = "Select * FROM  CLIENTEPLAN  ORDER BY ID"
			'response.write StrSql
			Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

			Do Until xgurs.eof


				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.' ID='***" & xgurs(0) & "' "

				If CStr(xgurs(0)) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If



				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If


				xgurs.movenext
			Loop

			Call ofv.cerrarconsulta(xgurs)

			If XGUITEMCD <> "NO" Then

				'CALL getUsrIngroup(XGUITEMCD)


				strsql = "Select * FROM  CLIENTEPLAN  WHERE ID=" & XGUITEMCD
				'  RESPONSE.WRITE STRSQL
				Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)
				Do While Not xgurs.EOF
					XGUTABKEY=Split(xgurs(0) ,Chr(9))
					XGUTABDAT=Split(xgurs(1)   ,Chr(9))

					xgurs.MOVENEXT
				Loop
				Call ofv.cerrarconsulta(xgurs)
				'XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
				'XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

				xTrColor = "#FFFFFF"

				WWWW = InStr(UCase(REQUEST.SERVERVARIABLES("SCRIPT_NAME") & ""),"AGREGARDIST.ASP")

				If WWWW > 0 Then

					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor='#6699CC' "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=VTI STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#000000;COLOR:#FFFFFF;' "
					XGUFBTD2 = XGUFBTD2 & " TITLE='SELECCIONAR TODOS LOS USUARIOS.'  ID='---XXXTODOSXXX" & XGUITEMCD & "' "

					If (XGUM4FL & "TODOS") = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & "XXXTODOSXXX" & XGUITEMCD & "','" & "TODOS LOS USUARIOS" & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & "TODOS LOS USUARIOS" & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"

				End If


				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR EMMPRESA.'  ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "FND','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If


			Case "G"

			XGUFILDSC1 = ":&nbsp;GRUPO"

			xTrColor = "#FFFFFF"

			If XGUITEMCD <> "NO" Then
				XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC1 = ""
			End If


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If

			If session("USRPRV")  = "I"  Or session("USRPRV")  = "E" Then
				strsql = "Select ID,DESCRIPCION From GRUPOS  order by descriPCION "  ' ok
				'strsql = "Select mapgrp,ID From GRUPOS  order by descriPCION " ' probar

			Else
				strsql = "Select ID,DESCRIPCION From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
			End If
			If request.querystring("OPWK") = "REASIGNAR" Then

				xxx = getGroupsUser(session("usuario"))

				grps = Split(session("SGU_grupos"),Chr(9))

				grupos = ""
				delim = ""

				For i=0 To UBound(grps)

					xgpU= Replace(Trim(grps(i) &""),"'","''")

					If xgpU <> "" Then
						grupos = grupos & delim & "'" & grps(i) & "'"
						delim = ","
					End If

				Next


				strsql = "select a.mapgrp,b.descripcion from gruposaprob a "
				strsql = strsql & " inner join grupos b on (a.mapgrp = b.id) "
				strsql = strsql & " where a.estado = 1 and a.mapgrp in (" & grupos & ")"
			End If


			Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)
			delim=""
			sgrupos=""


			Do While Not xgurs.eof
				sgrupos = sgrupos & delim & RTrim(LTrim(xgurs(0) & "" ))
				delim = Chr(9)
				xgurs.movenext
			Loop

			xx= buscarGrupoAd(sgrupos)


			XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
			XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

			xTrColor = "#FFFFFF"

			For XGUINX=0 To UBound(XGUTABKEY)



				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.' ID='***" & XGUTABKEY(XGUINX) & "' "

				If CStr(XGUTABKEY(XGUINX)) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If



				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUTABKEY(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & XGUTABDAT(XGUINX) & "</TD>"
				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If
			Next


			' end if

			'    xgurs.movenext
			'    loop

			'   call ofv.cerrarconsulta(xgurs)

			If XGUITEMCD <> "NO" Then

				Call getUsrIngroup(XGUITEMCD)

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				WWWW = InStr(UCase(REQUEST.SERVERVARIABLES("SCRIPT_NAME") & ""),"AGREGARDIST.ASP")

				If WWWW > 0 Then

					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor='#6699CC' "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=VTI STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#000000;COLOR:#FFFFFF;' "
					XGUFBTD2 = XGUFBTD2 & " TITLE='SELECCIONAR TODOS LOS USUARIOS.'  ID='---XXXTODOSXXX" & XGUITEMCD & "' "

					If (XGUM4FL & "TODOS") = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & "XXXTODOSXXX" & XGUITEMCD & "','" & "TODOS LOS USUARIOS" & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & "TODOS LOS USUARIOS" & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"

				End If


				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If



			Case "S"

			XGUFILDSC1 = ":&nbsp;SECTOR"

			xTrColor = "#FFFFFF"

			If XGUITEMCD <> "NO" Then
				XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC1 = ""
			End If


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If
			If session("ssec")="" Then session("ssec")=1

			'If session("ssec") = 1 Then

			'		strsql = "Select * From Eco where ecoun=" & session("ssec") & "order by descri"

			'	Else

			strsql = "Select * From Eco where ecoun=" & session("ssec") & " order by descri"

			'	End If

			Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

			Do Until xgurs.eof


				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

				If CStr(xgurs(0)) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If


				xgurs.movenext
			Loop

			Call ofv.cerrarconsulta(xgurs)


			If XGUITEMCD <> "NO" Then

				Call getUsrInSECT(XGUITEMCD,"")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If




			Case "L"

			XGUFILDSC1 = ":&nbsp;LETRA"

			xTrColor = "#F9F9F9"


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If


			Dim ATABXGULTR(28)

			ATABXGULTR(1)  = "A"
			ATABXGULTR(2)  = "B"
			ATABXGULTR(3)  = "C"
			ATABXGULTR(4)  = "D"
			ATABXGULTR(5)  = "E"
			ATABXGULTR(6)  = "F"
			ATABXGULTR(7)  = "G"
			ATABXGULTR(8)  = "H"
			ATABXGULTR(9)  = "I"
			ATABXGULTR(10) = "J"
			ATABXGULTR(11) = "K"
			ATABXGULTR(12) = "L"
			ATABXGULTR(13) = "M"
			ATABXGULTR(14) = "N"
			ATABXGULTR(15) = "O"
			ATABXGULTR(16) = "P"
			ATABXGULTR(17) = "Q"
			ATABXGULTR(18) = "R"
			ATABXGULTR(19) = "S"
			ATABXGULTR(20) = "T"
			ATABXGULTR(21) = "U"
			ATABXGULTR(22) = "V"
			ATABXGULTR(23) = "W"
			ATABXGULTR(24) = "X"
			ATABXGULTR(25) = "Y"
			ATABXGULTR(26) = "Z"
			ATABXGULTR(27) = "NO"
			ATABXGULTR(28) = "NO"


			For XGUINX=1 To UBound(ATABXGULTR) Step 4

				XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If ATABXGULTR(XGUINX) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " >"

				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & "</TD>"


				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If ATABXGULTR(XGUINX+1) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If

				XGUFBTD1 = XGUFBTD1 & " >"
				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
				XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
				XGUFBTD1 = XGUFBTD1 & "</TD>"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If  ATABXGULTR(XGUINX+2) = "NO" Then

					XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

				Else
					If ATABXGULTR(XGUINX+2) = XGUITEMCD Then
						XGUFBTD1 = XGUFBTD1 & XGUSTYLS
					Else
						XGUFBTD1 = XGUFBTD1 & XGUSTYLN
					End If
					XGUFBTD1 = XGUFBTD1 & " >"
					XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
					XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

				End If

				XGUFBTD1 = XGUFBTD1 & "</TD>"

				XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

				If  ATABXGULTR(XGUINX+3) = "NO" Then

					XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

				Else
					If ATABXGULTR(XGUINX+3) = XGUITEMCD Then
						XGUFBTD1 = XGUFBTD1 & XGUSTYLS
					Else
						XGUFBTD1 = XGUFBTD1 & XGUSTYLN
					End If
					XGUFBTD1 = XGUFBTD1 & " >"
					XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
					XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

				End If

				XGUFBTD1 = XGUFBTD1 & "</TD>"



				XGUFBTD1 = XGUFBTD1 & "</TR>"


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If

			Next




			If XGUITEMCD <> "NO" Then

				Call getUsrInTEXT(XGUITEMCD & "*")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If




			Case "T"

			XGUFILDSC1 = ":&nbsp;TEXTO"



			xTrColor = "#FFFFFF"


			If XGUCODIGO <> "NO" Then
				XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
			Else
				XGUFNC2 = ""
			End If



			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT ALIGN=CENTER TITLE='SELECCIONAR FILTRO.'  "

			If XGUITEMCD <> "NO" Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"
			XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE='" & XGUITEMCD & "' CLASS=DT size=35 >"

			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
			XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If XGUITEMCD <> "NO" Then

				Call getUsrInTEXT("*" & XGUITEMCD & "*")

				XGUTABKEY = Split(session("SGU_UIGuserid"),Chr(9))
				XGUTABDAT = Split(session("SGU_UIGdescri"),Chr(9))

				xTrColor = "#FFFFFF"

				For XGUINX=0 To UBound(XGUTABKEY)


					XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
					XGUFBTD2 = XGUFBTD2 & " >"

					XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

					If (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO Then
						XGUFBTD2 = XGUFBTD2 & XGUSTYLS
					Else
						XGUFBTD2 = XGUFBTD2 & XGUSTYLN
					End If

					XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
					XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
					XGUFBTD2 = XGUFBTD2 & "</TR>"


					If xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If


				Next




			End If


			Case Else

			XGUFILDSC1 = "&nbsp;"

			XGUFILTRO = "NO"
			XGUITEMCD = "NO"
			XGUCODIGO = "NO"
			XGUCODDSC = "NO"


		End Select






		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


		XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;padding: 5px 0px;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#dddddd;overflow:auto;border-top:none;border-radius: 0px 0px 10px 10px;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

		XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;padding: 5px 0px;"
		XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#dddddd;overflow:auto;border-top:none;border-radius: 0px 0px 10px 10px;'>"
		XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

		XGUCONTENT = ""
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=TTR style='border-radius:10px 10px 0px 0px;' width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
		If Session("EMPMUN")  = "S" And XGUFILTRO="TE" Then
			XGUCONTENT = XGUCONTENT & "<td class=TTR style='border-radius:10px 10px 0px 0px;' width=50% >&nbsp;EMPRESA</td>"
		Else
			XGUCONTENT = XGUCONTENT & "<td class=TTR style='border-radius:10px 10px 0px 0px;' width=50% >&nbsp;USUARIOS</td>"
		End If

		XGUCONTENT = XGUCONTENT & "</tr>"
		XGUCONTENT = XGUCONTENT & "<tr>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
		XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
		XGUCONTENT = XGUCONTENT & "</tr>"


	End Select


	Call ofv.cerrarconn(xgucn)


	Whtml = ""

	Whtml = Whtml & "<center>"
	'Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:650;"
	'Whtml = Whtml & "<DIV id='BSUSR' class='norma-card' style='Width:50%;'"
	Whtml = Whtml & "<DIV id='BSUSR' class='' style='Width:50%;'"
	Whtml = Whtml & "'>"
	Whtml = Whtml & "<CENTER>"

	Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=90% cellspacing=0 cellpadding=2 >"
	Whtml = Whtml & "<tr HEIGHT=25 >"
	Whtml = Whtml & "<td class=ttt style='border-radius: 10px 10px 0px 0px;' width=99%  >&nbsp;&nbsp;BUSCAR" & XGUFILDSC & "&nbsp;</td>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR0""  class='aaa' width=99%  cellspacing=0 cellpadding=0 "
	Whtml = Whtml & " style=""border-radius:0px;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"
	Whtml = Whtml & "<tD ALIGN=LEFT >"
	Whtml = Whtml & XGUCONTEMS
	Whtml = Whtml & "</tD>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR2""  class='aaa' width=99%  cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-radius:0px;border-top:none;border-bottom:none;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"

	Whtml = Whtml & XGUFBBOT5

	If session("APLMDCLI") <>  "GP" Then
		Whtml = Whtml & XGUFBBOT1
	End If

	Whtml = Whtml & XGUFBBOT2
	Whtml = Whtml & XGUFBBOT3
	Whtml = Whtml & XGUFBBOT4


	If dsDSC <> "NO" Then
		Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#dddddd;' >" & XGUCODDSC & "</td>"
	Else
		Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
	End If

	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR3""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-radius:0px 0px 10px 10px;"" "
	Whtml = Whtml & " >"

	Whtml = Whtml & XGUCONTENT

	Whtml = Whtml & "</table>"
	Whtml = Whtml & "</CENTER>"
	Whtml = Whtml & "</DIV>"
	Whtml = Whtml & "</center>"




	Whtml = Whtml & "<script language=javascript >"
	Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4) { " & Chr(13)
	' Whtml = Whtml & " alert(document.getElementById('XGUFRM'));  "
	Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU1; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU2; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUCODIGO').value=XGU3; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUCODDSC').value=XGU4; " & Chr(13)
	Whtml = Whtml & " } " & Chr(13)
	'Whtml = Whtml & XGUFNC3
	Whtml = Whtml & XGUFNC1
	Whtml = Whtml & XGUFNC2

	Whtml = Whtml & "</script>"

	Whtml = Whtml & "<FORM id='XGUFRM' NAME='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' id='XGUMASTER' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFILTRO' id='XGUFILTRO' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUITEMCD' id='XGUITEMCD' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODIGO' id='XGUCODIGO' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODDSC' id='XGUCODDSC' >"
	Whtml = Whtml & "</FORM>"



	SESSION("FBUSR") = XGUCODIGO
	SESSION("FBSLC") = XGUMASTER
	SESSION("XGUFILTRO")=XGUFILTRO


	GenBuscaUsr = Whtml


End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCIONES DE BUSQUEDa de Documentos ( documento-solapa-1er nivel)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GenBuscaObj(XGUCOBJI,XGUTOBJI,XGUN1I,XGUN2I,XGUN3I,XGUN4I,XGUN5I)


	XGUFBURL = request.servervariables("SCRIPT_NAME")


	XGUFBURL = Right(XGUFBURL,Len(XGUFBURL) - (InStrRev(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

	XGUMASTER = REQUEST.FORM("XGUMASTER")

	XGUITEMCD = REQUEST.FORM("XGUITEMCD")
	XGUFILTRO = REQUEST.FORM("XGUFILTRO")
	cmp=REQUEST.QUERYSTRING("cmp")
	'RESPONSE.WRITE XGUCOBJI & "  " &XGUTOBJI& "  " &XGUN1I& "  " &XGUN2I& "  " &XGUN3I& "  " &XGUN4I& "  " &XGUN5I
	xcodf=request.querystring("codobjeto")&""

	'response.write request.querystring

	If  XGUMASTER = "" Then
		If XGUCOBJI & "" = xcodf Then
			XGUFILTRO="11"
			XGUITEMCD = xcodf
		Else

			XGUFILTRO = "TX"
			XGUITEMCD = "NO"
		End If



		XGUMASTER = "0"

		XGUFCOBJ = "0"
		XGUFTOBJ = "0"
		XGUFNIV1 = "0"
		XGUFNIV2 = "0"
		XGUFNIV3 = "0"
		XGUFNIV4 = "0"
		XGUFNIV5 = "0"

		If   XGUCOBJI <> "0" Then
			XGUFCOBJ = XGUCOBJI
			XGUFTOBJ = XGUTOBJI
		End If

		If   XGUN1I <> "0" Then
			XGUFNIV1 = XGUN1I
			XGUMASTER = "1"
		End If

		If   XGUN2I <> "0" Then
			XGUFNIV2 = XGUN2I
			XGUMASTER = "2"
		End If

		If   XGUN3I <> "0" Then
			XGUFNIV3 = XGUN3I
			XGUMASTER = "3"
		End If

		If   XGUN4I <> "0" Then
			XGUFNIV4 = XGUN4I
			XGUMASTER = "4"
		End If

		If   XGUN5I <> "0" Then
			XGUFNIV5 = XGUN5I
			XGUMASTER = "5"
		End If

	Else

		XGUFCOBJ = REQUEST.FORM("XGUFCOBJ")
		XGUFTOBJ = REQUEST.FORM("XGUFTOBJ")
		XGUFNIV1 = REQUEST.FORM("XGUFNIV1")
		XGUFNIV2 = REQUEST.FORM("XGUFNIV2")
		XGUFNIV3 = REQUEST.FORM("XGUFNIV3")
		XGUFNIV4 = REQUEST.FORM("XGUFNIV4")
		XGUFNIV5 = REQUEST.FORM("XGUFNIV5")


	End If



	XGUCONTEMS = ""



	XGUFNC1 = ""
	XGUFNC2 = ""
	XGUFNC3 = ""


	Set xgucn = ofv.conectar(ofv.strconn2)

	XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#f7aa22;cursor:pointer;background-color:#bbddfe;' "
	XGUSTYLN = " STYLE='cursor:pointer;' "

	XGUFBUSR = XGUCODIGO

	XGUFBTAB1  = ""
	XGUFBTAB2  = ""
	XGUFBTD1   = ""
	XGUFBTD2   = ""

	XGUFBTDF0  = ""
	XGUFBTDF1  = ""
	XGUFBTDF2  = ""
	XGUFBTDF3  = ""
	XGUFBTDF4  = ""
	XGUFBTDF5  = ""
	XGUFBTDF6  = ""


	XGUFILDSC = "&nbsp;"

	XGUImG    = "<font face='corflaticon' size=3 color='#f7aa22'>4</font>"


	' -------------------------

	''XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TI','NO');document.getElementById('XGUFRM').submit();return false;"" >"
	'' "clactfrm('" & XGUMASTER & "','TI','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

	'' XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TODOS' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','A','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	'XGUMASTER="11"
	If XGUMASTER <> "11" And XGUFILTRO <> "11"   Then
		XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TIPO' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TI','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TL','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TX','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"



		'response.write session("un")

		If REQUEST.FORM("XGUITEMCD") =  ""  Or  Not IsNumeric(REQUEST.FORM("XGUITEMCD")) Then
			If session("SETIDEMPRE") = "1" Or session("SETIDEMPRE") = "" Then session("SETIDEMPRE")=  SESSION("D" & session("ncliente") & "-0" ) 'SESSION("CLIEMP")

			If session("SETIDEMPRE")="" Then session("SETIDEMPRE")=session("UN")
		Else
			session("SETIDEMPRE")=  REQUEST.FORM("XGUITEMCD")
		End If


		strsql = "Select CLIENTE From CLIENTEPLAN WHERE ID=" & session("SETIDEMPRE")
		'response.write strsql
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		If Not xgurs.EOF Then
			Scliente  = " : " & xgurs(0)
		End If
		Call OFV.CERRARCONSULTA(xgurs)

		XGUFBBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' STYLE='text-align:right' VALUE='EMPRESA" & Scliente & "'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TE','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

		XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
		XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
		XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
		XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"
		XGUFBBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI5 & "</FORM>"

	End If

	XGUFILDSC = "&nbsp;componentes"
	'RESPONSE.WRITE "<BR>eeeeeeeeeeeeeeeeeeeeeeee" & XGUFILTRO  & "<br>eeeeeeeeeeeeee "  & XGUITEMCD
	Select Case XGUFILTRO

		Case "NO"





		Case "TE"

		XGUFILDSC1 = ":&nbsp;EMPRESAS"

		xTrColor = "#FFFFFF"

		If XGUITEMCD <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If

		XEMPRE = Split(session("EmpHbl"),Chr(9))

		DELIM=""
		idempre = ""

		For i=0 To UBound(XEMPRE)

			idempre = idempre & DELIM & XEMPRE(i)
			DELIM = ","

		Next

		If idempre = "" Then idempre = session("CLIEMP")

		strsql = "Select ID,CLIENTE From CLIENTEPLAN where id in (" & idempre & ") order by ID "
		'response.write strsql & " strsql " & "<br>"
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='Seleccionar item " & xgurs(1) & ".' ID='***" & xgurs(0) & "' "

			If CStr(xgurs(0)) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			'''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TE','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TE','" & xgurs(0) & "');document.getElementById('XGUFRM').submit();return false;"" >"


			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTD1xxx = XGUFBTD1
		Case "TI"

		XGUFILDSC1 = ":&nbsp;TIPO COMPONENTE"

		xTrColor = "#FFFFFF"

		If XGUITEMCD <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If
		'   response.write "<br>" & SESSION("BASE") & "<br>" & cmp
		If SESSION("BASE")="K" Or cmp="K" Then
			strsql = "Select * From TIPOBJETO  where TIPOCPT='K' AND aplic IN ('EM', 'GLB') order by descri "
		Else
			strsql = "Select * From TIPOBJETO  where TIPOCPT NOT IN ('K','PD')  AND aplic IN ('EM', 'GLB') order by descri "
		End If

		'response.write strsql & " strsql " & "<br>"
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='Seleccionar item " & xgurs(1) & ".' ID='***" & xgurs(0) & "' "

			If CStr(xgurs(0)) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			'''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TI','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TI','" & xgurs(0) & "');document.getElementById('XGUFRM').submit();return false;"" >"


			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTD1xxx = XGUFBTD1






		Case "TL"

		XGUFILDSC1 = ":&nbsp;LETRA"

		xTrColor = "#F9F9F9"


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If


		Dim ATABXGULTR(28)

		ATABXGULTR(1)  = "A"
		ATABXGULTR(2)  = "B"
		ATABXGULTR(3)  = "C"
		ATABXGULTR(4)  = "D"
		ATABXGULTR(5)  = "E"
		ATABXGULTR(6)  = "F"
		ATABXGULTR(7)  = "G"
		ATABXGULTR(8)  = "H"
		ATABXGULTR(9)  = "I"
		ATABXGULTR(10) = "J"
		ATABXGULTR(11) = "K"
		ATABXGULTR(12) = "L"
		ATABXGULTR(13) = "M"
		ATABXGULTR(14) = "N"
		ATABXGULTR(15) = "O"
		ATABXGULTR(16) = "P"
		ATABXGULTR(17) = "Q"
		ATABXGULTR(18) = "R"
		ATABXGULTR(19) = "S"
		ATABXGULTR(20) = "T"
		ATABXGULTR(21) = "U"
		ATABXGULTR(22) = "V"
		ATABXGULTR(23) = "W"
		ATABXGULTR(24) = "X"
		ATABXGULTR(25) = "Y"
		ATABXGULTR(26) = "Z"
		ATABXGULTR(27) = "NO"
		ATABXGULTR(28) = "NO"


		For XGUINX=1 To UBound(ATABXGULTR) Step 4

			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If ATABXGULTR(XGUINX) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
			''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"





			XGUFBTD1 = XGUFBTD1 & "</TD>"


			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If ATABXGULTR(XGUINX+1) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"
			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
			''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX+1) & "');document.getElementById('XGUFRM').submit();return false;"" >"





			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If  ATABXGULTR(XGUINX+2) = "NO" Then

				XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

			Else
				If ATABXGULTR(XGUINX+2) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If
				XGUFBTD1 = XGUFBTD1 & " >"
				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
				''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



				XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX+2) & "');document.getElementById('XGUFRM').submit();return false;"" >"






			End If

			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If  ATABXGULTR(XGUINX+3) = "NO" Then

				XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

			Else
				If ATABXGULTR(XGUINX+3) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If
				XGUFBTD1 = XGUFBTD1 & " >"
				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
				'''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


				XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX+3) & "');document.getElementById('XGUFRM').submit();return false;"" >"




			End If

			XGUFBTD1 = XGUFBTD1 & "</TD>"



			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If

		Next









		Case "TX"

		XGUFILDSC1 = ":&nbsp;TEXTO"



		xTrColor = "#FFFFFF"


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
		XGUFBTD1 = XGUFBTD1 & " >"

		XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

		If XGUITEMCD <> "NO" Then
			XGUFBTD1 = XGUFBTD1 & XGUSTYLS
		Else
			XGUFBTD1 = XGUFBTD1 & XGUSTYLN
		End If

		XGUFBTD1 = XGUFBTD1 & " >"
		XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE='" & XGUITEMCD & "' CLASS=DT size=40 >&nbsp;&nbsp;"

		XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
		''    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


		XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TX','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TX',document.getElementById('XGUBTEXTO').value);document.getElementById('XGUFRM').submit();return false;"" >"




		XGUFBTD1 = XGUFBTD1 & "</TD>"

		XGUFBTD1 = XGUFBTD1 & "</TR>"



		Case "11"


		Case Else

		XGUFILDSC1 = "&nbsp;"

		XGUFILTRO = "11"
		XGUITEMCD = "11"
		XGUCODIGO = "NO"
		XGUCODDSC = "NO"


	End Select






	XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
	XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


	XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
	XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

	XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
	XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
	XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

	XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
	XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
	XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

	XGUCONTENT = ""
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
	XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;COMPONENTES</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"

	XGUCONTENTM = XGUFBDIV1

	' ------------------------

	'response.write XGUMASTER & "  XGUMASTER " & "<br>"
	'response.write "form=" & request.form  & "<br>"
	'response.write "QUERYSTRING=" & request.QUERYSTRING & "<br>"

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' ARMADO DE DIV SEGUN NIVEL DE PROFUNDIDAD ...
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Select Case XGUMASTER



		'  IF XGUCODIGO <> "NO" THEN
		'     XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
		'  ELSE
		'     XGUFNC2 = ""
		'  END IF


		Case "TEFND"

		XGUFNC3 = "clactfrm('TX','0','0','0','0','0','0','0','TX','NO');document.getElementById('XGUFRM').submit();"

		Case "TE"
		Call GETLISTOBJ(XGUMASTER,"0",XGUITEMCD)


		XGUTABKEY  = Split(XGUTABLISKEY,Chr(9))
		XGUTABKEY2 = Split(XGUTABLISKEY2,Chr(9))
		XGUTABDAT  = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"

		If XGUFCOBJ <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If


		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR EMPRESA.'  ID='---" & XGUTABKEY(XGUINX) & "' "

			If CStr(XGUTABKEY(XGUINX)) = CStr(XGUFCOBJ) Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If



			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('TEFND','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD & "');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next

		Case "TI","TL","TX"


		'RESPONSE.WRITE XGUMASTER &" "& XGUITEMCD

		Call GETLISTOBJ(XGUMASTER,"0",XGUITEMCD)


		XGUTABKEY  = Split(XGUTABLISKEY,Chr(9))
		XGUTABKEY2 = Split(XGUTABLISKEY2,Chr(9))
		XGUTABDAT  = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"



		If XGUFCOBJ <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR COMPONENTE.'  ID='---" & XGUTABKEY(XGUINX) & "' "

			If CStr(XGUTABKEY(XGUINX)) = CStr(XGUFCOBJ) Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If



			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next
		Case "0"

		' CALL GETLISTOBJ("L","0","0")
		Call GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
		XGUTABKEY  = Split(XGUTABLISKEY,Chr(9))
		XGUTABKEY2 = Split(XGUTABLISKEY2,Chr(9))
		XGUTABDAT  = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"



		If XGUFCOBJ <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR DOCUMENTO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

			If CStr(XGUTABKEY(XGUINX)) = CStr(XGUFCOBJ) Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If



			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next




		Case "1","11"


		Call GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT style='cursor:pointer;' "
		' XGUFBTDF0 = XGUFBTDF0 & " ONCLICKW=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0  & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLP("L",XGUFCOBJ,XGUFTOBJ,"0")
		XGUTABKEY = Split(XGUTABLISKEY,Chr(9))
		XGUTABDAT = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"


		If XGUFNIV1 <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFNIV1 & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR SOLAPA.'  ID='---" & XGUTABKEY(XGUINX) & "' "


			If CStr(XGUTABKEY(XGUINX)) = XGUFNIV1 Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If

			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUTABKEY(XGUINX) & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next



		Case "2"

		Call GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
		XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF1 = XGUFBTDF1 & "<TR  "
		XGUFBTDF1 = XGUFBTDF1 & " >"

		XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
		XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
		XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF1 = XGUFBTDF1 & "</TR>"





		Call GETLISTNV2("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,"0")
		XGUTABKEY = Split(XGUTABLISKEY,Chr(9))
		XGUTABDAT = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"


		If XGUFNIV2 <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFNIV2 & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If


		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 1ER. NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



			If CStr(XGUTABKEY(XGUINX)) = XGUFNIV2 Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If


			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUTABKEY(XGUINX) & "','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next





		Case "3"

		Call GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
		XGUFBTDF0 = XGUFBTDF0 & " ONCLICKX=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF1 = XGUFBTDF1 & "<TR  "
		XGUFBTDF1 = XGUFBTDF1 & " >"

		XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
		XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
		XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF1 = XGUFBTDF1 & "</TR>"



		Call GETLISTNV2("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF2 = XGUFBTDF2 & "<TR  "
		XGUFBTDF2 = XGUFBTDF2 & " >"

		XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
		XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
		XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF2 = XGUFBTDF2 & "</TR>"





		Call GETLISTNV3("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,"0")
		XGUTABKEY = Split(XGUTABLISKEY,Chr(9))
		XGUTABDAT = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"

		If XGUFNIV3 <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFNIV3 & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 2DO.NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



			If CStr(XGUTABKEY(XGUINX)) = XGUFNIV3 Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If


			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('4','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUTABKEY(XGUINX) & "','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next


		Case "4"

		Call GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
		XGUFBTDF0 = XGUFBTDF0 & " ONCLICKX=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF1 = XGUFBTDF1 & "<TR  "
		XGUFBTDF1 = XGUFBTDF1 & " >"

		XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
		XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
		XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF1 = XGUFBTDF1 & "</TR>"



		Call GETLISTNV2("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF2 = XGUFBTDF2 & "<TR  "
		XGUFBTDF2 = XGUFBTDF2 & " >"

		XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
		XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
		XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF2 = XGUFBTDF2 & "</TR>"



		Call GETLISTNV3("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,XGUFNIV3)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF3 = XGUFBTDF3 & "<TR  "
		XGUFBTDF3 = XGUFBTDF3 & " >"

		XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT  "
		XGUFBTDF3 = XGUFBTDF3 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUFNIV3 & "','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF3 = XGUFBTDF3 & XGUImG & "&nbsp;"
		XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF3 = XGUFBTDF3 & "</TR>"





		XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
		XGUFBTD2 = XGUFBTD2 & " >"
		XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI "
		XGUFBTD2 = XGUFBTD2 & XGUSTYLN
		XGUFBTD2 = XGUFBTD2 & "  >"
		XGUFBTD2 = XGUFBTD2 & "----------------------------------------------------------------------</TD>"
		XGUFBTD2 = XGUFBTD2 & "</TR>"






	End Select



	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



	'XGUMASTER

	If XGUMASTER > "0" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


	End If

	If XGUMASTER > "1" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


	End If

	If XGUMASTER > "2" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


	End If

	If XGUMASTER > "3" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


	End If

	If XGUMASTER > "0" Then

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

	End If




	XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
	XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

	XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
	XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
	XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

	XGUCONTENT = ""
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>" '& "#####" &  XGUMASTER & "####"


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	Call ofv.cerrarconn(xgucn)


	Whtml = ""

	Whtml = Whtml & "<center>"
	Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:650;"
	Whtml = Whtml & "'>"
	Whtml = Whtml & "<CENTER>"

	Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=99% cellspacing=0 cellpadding=2 >"
	Whtml = Whtml & "<tr HEIGHT=25 >"
	Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;&nbsp;SELECCIONAR&nbsp;DOCUMENTO&nbsp;Y&nbsp;CONTENIDO&nbsp;</td>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR9""  class='aaa' width=99%  cellspacing=0 cellpadding=0 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"
	Whtml = Whtml & "<tD ALIGN=RIGHT >"
	Whtml = Whtml & XGUCONTEMS
	Whtml = Whtml & "</tD>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR2""  class='aaa' width=99%  cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"
	If Session("EMPMUN")  = "S" Then
		' XGUFBBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESA : '" & XGUCODDSC  &  "  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','TE','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		'  XGUFBBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI5 & "</FORM>"
		Whtml = Whtml & XGUFBBOT5
	End If
	Whtml = Whtml & XGUFBBOT1
	Whtml = Whtml & XGUFBBOT2
	Whtml = Whtml & XGUFBBOT3
	Whtml = Whtml & XGUFBBOT4

	If XGUCODDSC <> "NO" Then
		Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
	Else
		Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
	End If

	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"
	Whtml = Whtml & "<table id=""MTBSUSR3""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"

	Whtml = Whtml & XGUCONTENT

	Whtml = Whtml & "</table>"
	Whtml = Whtml & "</CENTER>"
	Whtml = Whtml & "</DIV>"
	Whtml = Whtml & "</center>"' '& XGUITEMCD


	Whtml = Whtml & "<FORM NAME='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' id='XGUFRM' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' id='XGUMASTER' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFCOBJ' id='XGUFCOBJ'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFTOBJ' id='XGUFTOBJ'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV1' id='XGUFNIV1'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV2' id='XGUFNIV2'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV3' id='XGUFNIV3'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV4' id='XGUFNIV4'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV5' id='XGUFNIV5'>"

	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFILTRO' id='XGUFILTRO'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUITEMCD' id='XGUITEMCD'>"


	Whtml = Whtml & "</FORM>"

	Whtml = Whtml & "<script language=javascript >"
	Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4,XGU5,XGU6,XGU7,XGU8,XGU9) { " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFCOBJ').value=XGU1;  " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFTOBJ').value=XGU2;  " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV1').value=XGU3;  " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV2').value=XGU4;  " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV3').value=XGU5;  " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV4').value=XGU6;  " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV5').value=XGU7;  " & Chr(13)

	Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU8; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU9; " & Chr(13)


	Whtml = Whtml & " }" & Chr(13)
	Whtml = Whtml & XGUFNC3
	Whtml = Whtml & XGUFNC1
	Whtml = Whtml & XGUFNC2

	Whtml = Whtml & "</script>"




	GenBuscaObj = Whtml


End Function






Function GETLISTNV4(XGUFOPC, XGUFCOBJF, XGUFTOBJF, XGUFNIV1F, XGUFNIV2F, XGUFNIV3F, XGUFNIV4F)

	XGUTABLISKEY = "0"
	XGUTABLISDAT = ""





	If XGUFNIV1F = "9" Then

		Strsql = "SELECT * FROM ObjTabcon WHERE OBJID = 0" & XGUFNIV4F

		Set Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
		If Not Xrs.EOF Then
			XGUTABLISKEY = Xrs(0)

			xxxdescritmp = Xrs(6) & " " & Xrs(7) & ""
			If Len(xxxdescritmp) > 20 Then xxxdescritmp = Left(xxxdescritmp & "",20) & "..."
			XGUTABLISDAT = Xrs(4) & "." & xxxdescritmp

		End If
		Call ofv.cerrarconsulta(Xrs)

	End If

	If XGUFNIV1F = "11" Then

		Strsql = "SELECT * FROM OBJESTDET WHERE OBJID = 0" & XGUFNIV4F
		Set Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
		If Not Xrs.EOF Then
			XGUTABLISKEY = Xrs(0)

			xxxdescritmp = Xrs(8) & ""
			xxxdescrityp = UCase(Xrs(5) & "")

			If xxxdescrityp = "H" Then
				xxxdescritmp = "Imagen asociada."
			ElseIf xxxdescrityp = "I" Then
				xxxdescritmp = "Tabla asociada."
			ElseIf xxxdescrityp = "J" Then
				xxxdescritmp = "Componente asociado."
			ElseIf xxxdescrityp = "G" Then
				If Len(xxxdescritmp) > 25 Then xxxdescritmp = Left(xxxdescritmp & "",25) & "..."
			Else
				xxxdescritmp = Xrs(7) & ""
			End If

			XGUTABLISDAT = Xrs(4) & "." & xxxdescritmp
		End If
		Call ofv.cerrarconsulta(Xrs)

	End If



	GETLISTNV4 = "OK"

End Function

Function GETLISTOBJ(XGUFOPC,XGUFCOBJF,XGUFTOBJF)

	XGUTABLISKEY  = "0"
	XGUTABLISKEY2 = "NO"
	XGUTABLISDAT  = "SIN DEFINIR"

	XGUTABSEP    = Chr(9)
	'response.write XGUFOPC

	If XGUFOPC = "TE" Then
		If XGUFTOBJF  <> "NO" Then
			Strsql = "Select ID,CLIENTE From CLIENTEPLAN where ID =0" & session("SETIDEMPRE")
		Else
			Strsql = "Select ID,CLIENTE From CLIENTEPLAN where ID =0"' & session("SETIDEMPRE")
		End If
		'RESPONSE.WRITE STRSQL
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not Xrs2.eof Then

			XGUTABLISKEY  = Xrs2(0)
			XGUTABLISKEY2 = Xrs2(0)
			XGUTABLISDAT  = Xrs2(1)


		End If
		Call ofv.cerrarconsulta(Xrs2)

	ElseIf XGUFOPC = "E" Then


		'  Strsql = "Select * From csObjetosRDO where CODOBJETO = 0" & XGUFCOBJF
		If Session("CorApli") = "COREM20" Then
			Strsql = "SELECT  a.* FROM csobjetosRDOAPL a WHERE a.aplic IN ('EM', 'GLB') and  a.CODOBJETO = 0" & XGUFCOBJF

		Else
			Strsql = "Select * From csObjetosRDO where CODOBJETO = 0" & XGUFCOBJF
		End If

		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not Xrs2.eof Then

			XGUTABLISKEY  = Xrs2(0)
			XGUTABLISKEY2 = Xrs2(8)
			XGUTABLISDAT  = Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"


		End If
		Call ofv.cerrarconsulta(Xrs2)


	ElseIf XGUFOPC = "TI" Then



		Strsql = "Select * From csObjetosRDO a where a.estado = 3  OR a.ESTADO = 1 order by a.descri, a.OBJDESCRI"
		' response.write Strsql
		If Session("CorApli") = "COREM20" Then
			Strsql = "SELECT distinct TOP 100 a.* FROM csobjetosRDOAPL a "
			If Session("EMPMUN")  = "S" Then
				'STRSQL= STRSQL & " INNER JOIN OBJMEMPALC B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND B.CODEMPRE=" & SESSION("SETIDEMPRE") & ")"
				STRSQL= STRSQL & " INNER JOIN objetos B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND objun=" & SESSION("SETIDEMPRE") & ")"
			End If
			If SESSION("BASE") <> "K" And cmp <> "K" Then
				strsql=strsql & " INNER JOIN  TIPOBJETO ON (A.TIPOBJETO =TIPOBJETO.TIPOBJETO AND TIPOCPT not in  ('K','PD'))"
			Else
				strsql=strsql & " INNER JOIN  TIPOBJETO ON (A.TIPOBJETO =TIPOBJETO.TIPOBJETO AND TIPOCPT = 'K')"
			End If
			strsql= strsql & " WHERE a.aplic IN ('EM', 'GLB') and (a.estado = 3  OR a.ESTADO = 1) and a.tipobjeto='" & XGUFTOBJF  & "' "
			strsql=strsql &  " order by a.descri, a.OBJDESCRI"
		Else
			Strsql = "Select distinct TOP 100 a.* From csObjetosRDO a "
			If Session("EMPMUN")  = "S" Then
				'STRSQL= STRSQL & " INNER JOIN OBJMEMPALC B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND B.CODEMPRE=" & SESSION("SETIDEMPRE") & ")"
				STRSQL= STRSQL & " INNER JOIN objetos B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND objun=" & SESSION("SETIDEMPRE") & ")"
			End If
			strsql=strsql & " where (a.estado = 3  OR a.ESTADO = 1) and a.tipobjeto='" & XGUFTOBJF & "' "
			strsql=strsql &  " order by a.descri, a.OBJDESCRI"
		End If
		'response.write strsql
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do While Not Xrs2.eof

			XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
			XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(8)
			XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"



			Xrs2.movenext
		Loop


	ElseIf XGUFOPC = "TX" Then

		' Strsql = "Select * From csObjetosRDO a where a.estado = 3  OR a.ESTADO = 1 order by a.descri, a.OBJDESCRI"
		If Session("CorApli") = "COREM20" Then
			Strsql = "SELECT distinct TOP 100 A.* FROM csobjetosRDOAPL a  "
			If Session("EMPMUN")  = "S" Then
				'STRSQL= STRSQL & " INNER JOIN OBJMEMPALC B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND B.CODEMPRE=" & SESSION("SETIDEMPRE") & ")"
				STRSQL= STRSQL & " INNER JOIN objetos B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND objun=" & SESSION("SETIDEMPRE") & ")"
			End If
			If SESSION("BASE") <> "K" And cmp <> "K" Then
				strsql=strsql & " INNER JOIN  TIPOBJETO ON (A.TIPOBJETO =TIPOBJETO.TIPOBJETO AND TIPOCPT not in  ('K','PD'))"
			Else
				strsql=strsql & " INNER JOIN  TIPOBJETO ON (A.TIPOBJETO =TIPOBJETO.TIPOBJETO AND TIPOCPT = 'K')"
			End If
			STRSQL= STRSQL & " WHERE a.aplic IN ('EM', 'GLB') and (a.estado = 3  OR a.ESTADO = 1) and a.OBJDescri like('%" & XGUFTOBJF & "%')  order by a.descri, a.OBJDESCRI "
		Else
			Strsql = "Select distinct TOP 100 * From csObjetosRDO a where (a.estado = 3  OR a.ESTADO = 1) and a.OBJDescri like('%" & XGUFTOBJF & "%')   order by a.descri, a.OBJDESCRI"
		End If
		'RESPONSE.WRITE STRSQL
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do While Not Xrs2.eof

			XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
			XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(8)
			XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"



			Xrs2.movenext
		Loop

	ElseIf XGUFOPC = "TL" Then

		'Strsql = "Select * From csObjetosRDO where estado = 3  OR ESTADO = 1 order by descri, OBJDESCRI"
		If Session("CorApli") = "COREM20" Then
			Strsql = "SELECT distinct TOP 100 a.* FROM csobjetosRDOAPL a "
			If Session("EMPMUN")  = "S" Then
				'STRSQL= STRSQL & " INNER JOIN OBJMEMPALC B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND B.CODEMPRE=" & SESSION("SETIDEMPRE") & ")"
				STRSQL= STRSQL & " INNER JOIN objetos B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND objun=" & SESSION("SETIDEMPRE") & ")"
			End If
			If SESSION("BASE") <> "K" And cmp <> "K" Then
				strsql=strsql & " INNER JOIN  TIPOBJETO ON (A.TIPOBJETO =TIPOBJETO.TIPOBJETO AND TIPOCPT not in  ('K','PD'))"
			Else
				strsql=strsql & " INNER JOIN  TIPOBJETO ON (A.TIPOBJETO =TIPOBJETO.TIPOBJETO AND TIPOCPT = 'K')"
			End If
			STRSQL= STRSQL & "  WHERE a.aplic IN ('EM', 'GLB') and (a.estado = 3  OR a.ESTADO = 1) and a.OBJDescri like('" & XGUFTOBJF & "%')  order by a.descri, a.OBJDESCRI "
		Else
			Strsql = "Select distinct TOP 100 * From csObjetosRDO a where (a.estado = 3  OR a.ESTADO = 1) and a.OBJDescri like('" & XGUFTOBJF & "%')  order by a.descri, a.OBJDESCRI"
		End If
		'	response.write strsql
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do While Not Xrs2.eof

			XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
			XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(8)
			XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"



			Xrs2.movenext
		Loop


		Call ofv.cerrarconsulta(Xrs2)

	End If

	GETLISTOBJ = "OK"


End Function



Function GETLISTSLP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F)

	XGUTABLISKEY = "0"
	XGUTABLISDAT = "TODO EL DOCUMENTO"

	XGUTABSEP    = Chr(9)

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''' BUSCA LAS SOLAPAS DEL DOCUMENTO
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Dim stabla(20)
	Dim solord(20)
	Dim Nombres(20)
	Dim Solapas(20)

	solord(0) = 0
	solord(1) = 15
	solord(2) = 1
	solord(3) = 6
	solord(4) = 3
	solord(5) = 2
	solord(6) = 4
	solord(7) = 5
	solord(8) = 7
	solord(9) = 8
	solord(10) = 11
	solord(11) = 10
	solord(12) = 9
	solord(13) = 12
	solord(14) = 13
	solord(15) = 14
	solord(16) = 16
	solord(17) = 17
	solord(18) = 18
	solord(19) = 19
	solord(20) = 20


	For xx = 1 To 20
		solord(xx) = 17
		Nombres(xx) = "N"
	Next


	Strsql = "Select * From Objperfildoc Where codobjeto =" & XGUFCOBJF & " and tipobjeto ='" & XGUFTOBJF & "' "
	'response.write Strsql & " strsql 1 " & "<br>"
	Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Xrs.eof Then
		perfilsp = "N"
	End If

	If perfilsp = "N" Then
		Call ofv.cerrarconsulta(Xrs)

		sPerfil = "OBJMAIN"

		Strsql = "Select PERFIL From TIPOBJETO Where tipoBJETO='" & XGUFTOBJF & "'"
		'response.write Strsql & " strsql 2 " & "<br>"
		Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not Xrs.eof Then
			sPerfil = XRS(0)
		End If

		Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
		'response.write Strsql & " strsql 3 " & "<br>"
		Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	End If

	''response.write strsql & "<br>"


	For xx = 2 To 16
		stabla (xx - 1) = Xrs(xx + 30)
	Next


	yy = 0
	For xx = 1 To 15

		If CInt(stabla (xx)) < 99 Then
			yy = yy + 1
			solord (stabla (xx)) = xx
		End If
	Next
	tyy = yy



	For z=0 To 20
		Solapas(z) = "N"
	Next


	X=0
	For z = 2 To 16

		If Xrs(z) Then

			'''if (z <> 2 and z <> 4 and z <> 5 )  then
			If (z <> 2 And z <> 5 )  Then

				x = x + 1
				Solapas(z-1)  = XRS(z + 15)
				'response.write Solapas(z-1) & " " & z-1 & " Solapas (z) " & "<br>"
				If  z = 6 Then

					Solapas(16)  = "CIRCUITO COMPLETO"

					For xiord=0 To 19

						If solord(xiord) = 5 Then

							xmovey = 16

							For xiord2=xiord + 1 To 19
								xmoven           = solord(xiord2)
								solord(xiord2)   = xmovey
								xmovey           = xmoven
							Next



							Exit For
						End If
					Next


				End If

			End If
		End If

	Next


	Call ofv.cerrarconsulta(Xrs)

	xfinal = UBound(Solapas)






	If XGUFOPC = "E" Then

		Strsql = "Select * From csGenObjetos where CODOBJETO = 0" & XGUFCOBJF
		'response.write Strsql & " strsql " & "<br>"
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not Xrs2.eof Then

			XGUTABLISKEY = Xrs2(0)
			XGUTABLISDAT = Xrs2(1)


		End If
		Call ofv.cerrarconsulta(Xrs2)

		For x = xstrsol To xfinal
			If solapas(solord(x)) <> "N" Then
				'   response.write solord(x)
				If CLng(XGUFNIV1F) = CLng(solord(x)) Then

					XGUTABLISKEY = solord(x)
					XGUTABLISDAT = solapas(solord(x))

					Exit For

				End If
			End If
		Next



	Else


		For x = xstrsol To xfinal
			If solapas(solord(x)) <> "N" Then

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & solord(x)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & solapas(solord(x))

			End If
		Next




	End If

	GETLISTSLP = "OK"

End Function



Function GETLISTNV2(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F)

	XGUTABLISKEY = "0"
	XGUTABLISDAT = "TODO EL CONTENIDO "

	XGUTABSEP    = Chr(9)

	XGUTABLAF    = "NO"


	Select Case XGUFNIV1F

		Case "2"   'CONSIDERACIONES

		XGUTABLAF    = "OBJNWEDT"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM OBJNWEDT WHERE OBJID = 0" & XGUFNIV2F
			'response.write Strsql & " strsql 4818 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(3)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM OBJNWEDT WHERE codobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " order by consecuen "
			'response.write Strsql & " strsql 4833 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If

		'' NUEVO POR VERSION 4.6 - 25-02-2015

		Case "3"   'SECTORES

		'response.write " entro a Sectores - linea 4851 " & "<br>"

		XGUTABLAF    = "OBJSECTOR"

		If XGUFOPC = "E" Then

			Strsql = "SELECT OBJSECTOR.*, ECO.descri FROM OBJSECTOR, ECO "
			Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and eco.estado=3 and codobjeto='" & XGUFNIV2F & "' "
			Strsql = Strsql & " And (calificador = 'R' or calificador = 'C' or calificador = 'O') "
			'response.write Strsql & " strsql 4863 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(13)

			End If
			Call ofv.cerrarconsulta(Xrs)

		Else

			Strsql = "SELECT OBJSECTOR.*, ECO.descri FROM OBJSECTOR, ECO "
			Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and eco.estado=3 and codobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " And (calificador = 'R' or calificador = 'C' or calificador = 'O') "
			'response.write Strsql & " strsql 4877 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(13)

				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If

		'' NUEVO POR VERSION 4.6 - 25-02-2015 FIN


		Case "5"   'PROCEDIMIENTOS

		XGUTABLAF    = "CSCEventoSectorPerfil"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM CSCEventoSectorPerfil WHERE CODPROCED = '" & XGUFNIV2F & "' "
			'response.write Strsql & " strsql 4902 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(2) & "."  & Xrs(6) & ".&nbsp;" & Xrs(4) & "&nbsp;-&nbsp;"  & Xrs(3) & "&nbsp;-&nbsp;" & Xrs(5)


			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM CSCEventoSectorPerfil WHERE obcodobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " order by SECUEN, SECUENCIA "
			'response.write Strsql & " strsql 4918 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(2) & "."  & Xrs(6) & ".&nbsp;" & Xrs(4) & "&nbsp;-&nbsp;"  & Xrs(3) & "&nbsp;-&nbsp;" & Xrs(5)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If







		Case "7"   'CONTROLES

		XGUTABLAF    = "cscontrolsel"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM cscontrolsel WHERE OBJID = 0" & XGUFNIV2F
			'response.write Strsql & " strsql 4976 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(1)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM cscontrolsel WHERE codobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " order by DESCRI "
			'response.write Strsql & " strsql 4991 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(1)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If





		Case "8"   'ESQUEMAS

		XGUTABLAF    = "csesquemasel"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM csesquemasel WHERE OBJID = 0" & XGUFNIV2F
			'response.write Strsql & " strsql 4993 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(1)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM csesquemasel WHERE codobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " order by DESCRI "
			'response.write Strsql & " strsql 5008 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(1)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If






		Case "9"   'TABLAS

		XGUTABLAF    = "ObjTabDef"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM ObjTabDef WHERE OBJID = 0" & XGUFNIV2F
			'response.write Strsql & " strsql 5038 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(4)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM ObjTabDef WHERE codobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " order by tabsecuen "
			'response.write Strsql & " strsql 5053 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(4)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If

		Case "10" 'documentos

		Case "13" 'contratos

		XGUTABLAF = "Objhtm"

		If XGUFOPC = "E" Then

			Strsql = "SELECT * FROM Objhtm WHERE OBJID = 0" & XGUFNIV2F
			'response.write Strsql & " strsql 5077 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If Not Xrs.EOF Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(3) & "." & Xrs(4)

			End If
			Call ofv.cerrarconsulta(Xrs)

		Else

			Strsql = "SELECT * FROM Objhtm WHERE codobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " order by secuen "
			'response.write Strsql & " strsql 5090 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

			Do While Not Xrs.EOF

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP &  Xrs(3) & "." & Xrs(4)

				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If




		Case "11"   'CONTENIDO

		XGUTABLAF    = "OBJESTR"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM OBJESTR WHERE OBJID = 0" & XGUFNIV2F
			'response.write Strsql & " strsql 5073 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(3)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM OBJESTR WHERE codobjeto = 0" & XGUFCOBJF
			Strsql = Strsql & " order by CONsecuen "
			'response.write Strsql & " strsql 5088 " & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If



		Case Else

		XGUTABLAF    = "NO"

		'response.write " salida por CASE ELSE strsql 5111 " & "<br>"

	End Select




	GETLISTNV2 = "OK"

End Function


Function GETLISTNV3(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F,XGUFNIV3F)

	XGUTABLISKEY = "0"
	XGUTABLISDAT = "TODO EL CONTENIDO "

	XGUTABSEP    = Chr(9)




	If XGUFOPC = "E" Then

		If  XGUFNIV1F = "9" Then


			Strsql = "SELECT * FROM ObjTabblk WHERE OBJID = 0" & XGUFNIV3F

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(5)

			End If
			Call ofv.cerrarconsulta(Xrs)


		End If

		If  XGUFNIV1F = "11" Then

			Strsql = "SELECT * FROM OBJESTDET WHERE OBJID = 0" & XGUFNIV3F
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = UCase(Xrs(7))
			End If
			Call ofv.cerrarconsulta(Xrs)

		End If

	Else

		If  XGUFNIV1F = "9" Then


			Strsql = "SELECT * FROM ObjTabblk WHERE tabdef = 0" & XGUFNIV2F
			Strsql = Strsql & " order by tabblksec "

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(5)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)

		End If


		If  XGUFNIV1F = "11" Then

			strsql = " SELECT * FROM CSESTDET "
			strsql = strsql & " where detcon = " & XGUFNIV2F & "  "
			strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
			strsql = strsql & " and not ( "
			strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
			strsql = strsql & " ) and stmodi = '' AND DETTIPO < 'G' "
			strsql = strsql & " order by TSEC"

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do While Not Xrs.eof

				XGUINDN = "-&nbsp;"

				Select Case XRS(5)

					Case "B"

					XGUINDN = "-&nbsp;&nbsp;&nbsp;"


					Case "C"

					XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


					Case "D"

					XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


					Case "E"

					XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


					Case "F"

					XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"



				End Select

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & XGUINDN & UCase(Xrs(7))

				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If


	End If

	GETLISTNV3 = "OK"

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BUSCADOR DE NIVELES DE COMPONENTES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function GETLISTCMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF)

	XGUTABLISKEY  = "0"
	XGUTABLISKEY2 = "NO"
	XGUTABLISDAT  = "SIN DEFINIR"

	XGUTABSEP    = Chr(9)

	If XGUFOPC = "E" Then

		Strsql = "Select * From csacomple where CODCOMP = 0" & XGUFCOBJF
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not Xrs2.eof Then

			XGUTABLISKEY  = Xrs2(0)
			XGUTABLISKEY2 = Xrs2(1)
			XGUTABLISDAT  = Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)


		End If
		Call ofv.cerrarconsulta(Xrs2)



	ElseIf XGUFOPC = "TI" Then

		Strsql = "Select * From csacomple where compun= '" & session("SETIDEMPRE") & "' and tipoCOMP = '" & XGUFTOBJF & "' order by descri, TIPODESCRI"
		'response.write strsql
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do While Not Xrs2.eof

			XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
			XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
			XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



			Xrs2.movenext
		Loop
		Call ofv.cerrarconsulta(Xrs2)

	ElseIf XGUFOPC = "TX" Then

		Strsql = "Select * From csacomple where compun= '" & session("SETIDEMPRE") & "' and  descri like('%" & XGUFTOBJF & "%') order by descri, TIPODESCRI"
		'''    response.write "por texto: " & strsql & " <br>"
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do While Not Xrs2.eof

			XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
			XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
			XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



			Xrs2.movenext
		Loop
		Call ofv.cerrarconsulta(Xrs2)

	ElseIf XGUFOPC = "TL" Then

		Strsql = "Select * From csacomple where compun= '" & session("SETIDEMPRE") & "' and descri like('" & XGUFTOBJF & "%') order by descri, TIPODESCRI"
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do While Not Xrs2.eof

			XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
			XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
			XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



			Xrs2.movenext
		Loop
		Call ofv.cerrarconsulta(Xrs2)


	Else

		Strsql = "Select * From csacomple  where compun= '" & session("SETIDEMPRE") & "' and order by descri, TIPODESCRI"

		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do While Not Xrs2.eof

			XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
			XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
			XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



			Xrs2.movenext
		Loop
		Call ofv.cerrarconsulta(Xrs2)

	End If

	GETLISTCMP = "OK"


End Function







Function GETLISTSLPCMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F)

	XGUTABLISKEY = "0"
	XGUTABLISDAT = "TODO EL COMPONENTE"

	XGUTABSEP    = Chr(9)

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''' BUSCA LAS SOLAPAS DEL COMPONENTE
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Dim stabla(20)
	Dim solord(20)
	Dim Nombres(20)


	solord(0) = 0
	solord(1) = 15
	solord(2) = 1
	solord(3) = 6
	solord(4) = 3
	solord(5) = 2
	solord(6) = 4
	solord(7) = 5
	solord(8) = 7
	solord(9) = 8
	solord(10) = 11
	solord(11) = 10
	solord(12) = 9
	solord(13) = 12
	solord(14) = 13
	solord(15) = 14
	solord(16) = 16
	solord(17) = 17
	solord(18) = 18
	solord(19) = 19
	solord(20) = 20



	'''''''''''''''''''''''

	Dim Solapas(20)


	For z=0 To 20
		Solapas(z) = "N"
	Next


	Strsql = "Select * From tipocomp Where tipo = '" & XGUFTOBJF & "' "
	Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not Xrs.eof Then
		sPerfil  = Xrs(4)
	End If

	Call ofv.cerrarconsulta(Xrs)



	Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
	'' response.write strsql & "<br>"
	Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not Xrs.eof Then



		For xx = 1 To 20
			solord(xx) = 17
			Nombres(xx) = "N"
		Next



		For xx = 2 To 16
			stabla (xx - 1) = Xrs(xx + 30)
		Next



		yy = 0
		For xx = 1 To 15

			If CInt(stabla (xx)) < 99 Then
				yy = yy + 1
				solord (stabla (xx)) = xx
			End If
		Next
		tyy = yy

		xstrsol     = 1



		X=0
		For z = 2 To 16

			If Xrs(z) Then

				If  z <> 21 And z <> 4 And z <> 61 Then
					x = x + 1
					Solapas(z-1)  = Replace(XRS(z + 15)," ","*")

				End If
			End If

		Next

	End If

	Call ofv.cerrarconsulta(Xrs)


	xfinal = UBound(Solapas)


	If XGUFOPC = "E" Then

		Strsql = "Select * From CSACOMPLE where CODCOMP = 0" & XGUFCOBJF
		Set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not Xrs2.eof Then

			XGUTABLISKEY = Xrs2(0)
			XGUTABLISDAT = Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)


		End If
		Call ofv.cerrarconsulta(Xrs2)

		For x = xstrsol To xfinal
			If solapas(solord(x)) <> "N" Then
				'   response.write solord(x)
				If CLng(XGUFNIV1F) = CLng(solord(x)) Then

					XGUTABLISKEY = solord(x)
					XGUTABLISDAT = solapas(solord(x))

					Exit For

				End If
			End If
		Next



	Else


		For x = xstrsol To xfinal
			If solapas(solord(x)) <> "N" Then

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & solord(x)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP  & solord(x) & "&nbsp;" & solapas(solord(x))

			End If
		Next




	End If

	GETLISTSLPCMP = "OK"

End Function


Function GETLISTNV2CMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F)

	XGUTABLISKEY = "0"
	XGUTABLISDAT = "TODO EL CONTENIDO " & XGUFNIV1F

	XGUTABSEP    = Chr(9)

	XGUTABLAF    = "NO"


	Select Case XGUFNIV1F

		Case "2"   'estructura

		XGUTABLAF    = "COMPESTR"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(3)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
			Strsql = Strsql & " order by OBJID "

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If







		Case "4"   'imagenes

		XGUTABLAF    = "COMPIMAG"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(4)

				If XGUTABLISDAT = "" Then XGUTABLISDAT = Xrs(3)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
			Strsql = Strsql & " order by SECUENCIA, OBJID "
			''   response.write strsql & "<br>"
			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)

				WKXGUTABLISDAT = Xrs(4)

				If WKXGUTABLISDAT = "" Then WKXGUTABLISDAT = Xrs(3)

				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & WKXGUTABLISDAT



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If







		Case "6"   'archivos

		XGUTABLAF    = "COMPARC"

		If XGUFOPC = "E" Then


			Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not Xrs.eof Then
				XGUTABLISKEY = Xrs(0)
				XGUTABLISDAT = Xrs(4)

				If XGUTABLISDAT = "" Then XGUTABLISDAT = Xrs(3)

			End If
			Call ofv.cerrarconsulta(Xrs)


		Else


			Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
			Strsql = Strsql & " order by SECUENCIA, OBJID "

			Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not Xrs.eof

				XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)

				WKXGUTABLISDAT = Xrs(4)

				If WKXGUTABLISDAT = "" Then WKXGUTABLISDAT = Xrs(3)

				XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & WKXGUTABLISDAT



				Xrs.movenext
			Loop
			Call ofv.cerrarconsulta(Xrs)


		End If








		Case Else

		XGUTABLAF    = "NO"



	End Select




	GETLISTNV2CMP = "OK"

End Function


Function GETLISTNV3CMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F,XGUFNIV3F)

	XGUTABLISKEY = "0"
	XGUTABLISDAT = "TODO EL CONTENIDO "

	XGUTABSEP    = Chr(9)

	GETLISTNV3CMP = "OK"

End Function

Function GenBuscaCMP(XGUCOBJI,XGUTOBJI,XGUN1I,XGUN2I,XGUN3I,XGUN4I,XGUN5I)


	XGUFBURL = request.servervariables("SCRIPT_NAME")

	XGUFBURL = Right(XGUFBURL,Len(XGUFBURL) - (InStrRev(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

	XGUMASTER = REQUEST.FORM("XGUMASTER")
	XGUITEMCD = REQUEST.FORM("XGUITEMCD")
	XGUFILTRO = REQUEST.FORM("XGUFILTRO")

	XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#F7AA23;cursor:pointer;' "
	XGUSTYLN = " STYLE='cursor:pointer;' "


	'''IF  XGUMASTER = "TX" THEN XGUITEMCD = REQUEST.FORM("XGUBTEXTO")


	If  XGUMASTER = "" Then

		XGUMASTER = "FT"
		XGUFILTRO = "TI"
		XGUITEMCD = "NO"

		XGUFCOBJ = "0"
		XGUFTOBJ = "0"
		XGUFNIV1 = "0"
		XGUFNIV2 = "0"
		XGUFNIV3 = "0"
		XGUFNIV4 = "0"
		XGUFNIV5 = "0"


		If   XGUCOBJI <> "0" Then
			XGUFCOBJ = XGUCOBJI
			XGUFTOBJ = XGUTOBJI
		End If

		If   XGUN1I <> "0" Then
			XGUFNIV1 = XGUN1I
			XGUMASTER = "1"
		End If

		If   XGUN2I <> "0" Then
			XGUFNIV2 = XGUN2I
			XGUMASTER = "2"
		End If

		If   XGUN3I <> "0" Then
			XGUFNIV3 = XGUN3I
			XGUMASTER = "3"
		End If

		If   XGUN4I <> "0" Then
			XGUFNIV4 = XGUN4I
			XGUMASTER = "4"
		End If

		If   XGUN5I <> "0" Then
			XGUFNIV5 = XGUN5I
			XGUMASTER = "5"
		End If


	Else

		XGUFCOBJ = REQUEST.FORM("XGUFCOBJ")
		XGUFTOBJ = REQUEST.FORM("XGUFTOBJ")
		XGUFNIV1 = REQUEST.FORM("XGUFNIV1")
		XGUFNIV2 = REQUEST.FORM("XGUFNIV2")
		XGUFNIV3 = REQUEST.FORM("XGUFNIV3")
		XGUFNIV4 = REQUEST.FORM("XGUFNIV4")
		XGUFNIV5 = REQUEST.FORM("XGUFNIV5")

		'response.write XGUFCOBJ & " XGUFCOBJ " & "<br>"
		'response.write XGUFTOBJ & " XGUFTOBJ " & "<br>"

	End If



	XGUCONTEMS = ""



	XGUFNC1 = ""
	XGUFNC2 = ""
	' XGUFNC3 = ""


	Set xgucn = ofv.conectar(ofv.strconn2)

	XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#f7aa22;cursor:pointer;background-color:#bbddfe;' "
	XGUSTYLN = " STYLE='cursor:pointer;' "

	XGUFBUSR = XGUCODIGO

	XGUFBTAB1  = ""
	XGUFBTAB2  = ""
	XGUFBTD1   = ""
	XGUFBTD2   = ""

	XGUFBTDF0  = ""
	XGUFBTDF1  = ""
	XGUFBTDF2  = ""
	XGUFBTDF3  = ""
	XGUFBTDF4  = ""
	XGUFBTDF5  = ""


	XGUFILDSC = "&nbsp;"

	XGUImG    = "<font face='Corflaticon' size=3 color='#f7aa22'>4</font>"
	'response.write REQUEST.FORM & "  1<br>"
	'RESPONSE.WRITE  XGUMASTER & " 22"
	'If REQUEST.FORM("XGUITEMCD") =  ""  Or  Not IsNumeric(REQUEST.FORM("XGUITEMCD")) Then
	'	If REQUEST.FORM("XGUITEMCD") =  "" OR  REQUEST.FORM("XGUITEMCD")="NO" THEN
	If  XGUMASTER <> "TE" And session("SETIDEMPRE") = "1" Or session("SETIDEMPRE") = "" Then session("SETIDEMPRE")=  SESSION("D" & session("ncliente") & "-0" ) 'SESSION("CLIEMP")
	If XGUMASTER="TE" And XGUITEMCD <> "NO" Then session("SETIDEMPRE")=  REQUEST.FORM("XGUITEMCD")

	'	RESPONSE.WRITE  XGUMASTER & "22"
	'Else
	'	session("SETIDEMPRE")=  REQUEST.FORM("XGUITEMCD")
	'End If

	'response.write "SITE=" & session("SETIDEMPRE")
	If Session("EMPMUN") <> "S" Then session("SETIDEMPRE")="1"
	strsql = "Select CLIENTE From CLIENTEPLAN WHERE ID=" & session("SETIDEMPRE")

	Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

	If Not xgurs.EOF Then
		Scliente  = " : " & xgurs(0)
	End If
	Call OFV.CERRARCONSULTA(xgurs)


	''XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TI','NO');document.getElementById('XGUFRM').submit();return false;"" >"
	'' "clactfrm('" & XGUMASTER & "','TI','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	'' XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TODOS' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','A','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

	XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TIPO' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TI','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TL','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TX','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	XGUFBBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' STYLE='text-align:right' VALUE='EMPRESA" & Scliente & "'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TE','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
	XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
	XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
	XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"
	XGUFBBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI5 & "</FORM>"


	XGUFILDSC = "&nbsp;componentes"

	Select Case XGUFILTRO

		Case "NO"
		Case "TE"

		XGUFILDSC1 = ":&nbsp;EMPRESAS"

		xTrColor = "#FFFFFF"

		If XGUITEMCD <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If

		XEMPRE = Split(session("EmpHbl"),Chr(9))

		DELIM=""
		idempre = ""

		For i=0 To UBound(XEMPRE)

			idempre = idempre & DELIM & XEMPRE(i)
			DELIM = ","

		Next

		If idempre = "" Then idempre = session("CLIEMP")

		strsql = "Select ID,CLIENTE From CLIENTEPLAN where id in (" & idempre & ") order by ID "
		'response.write strsql & " strsql " & "<br>"
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='Seleccionar item " & xgurs(1) & ".' ID='***" & xgurs(0) & "' "

			If CStr(xgurs(0)) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			'''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TE','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TE','" & xgurs(0) & "');document.getElementById('XGUFRM').submit();return false;"" >"


			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTD1xxx = XGUFBTD1

		Case "TI"

		XGUFILDSC1 = ":&nbsp;TIPO COMPONENTE"

		xTrColor = "#FFFFFF"

		If XGUITEMCD <> "NO" Then
			XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC1 = ""
		End If


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If


		strsql = "Select * From TIPOCOMP  order by descri "
		''response.write strsql & " strsql " & "<br>"
		Set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

		Do Until xgurs.eof


			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='Seleccionar item " & xgurs(1) & ".' ID='***" & xgurs(0) & "' "

			If CStr(xgurs(0)) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If



			'''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TI','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & xgurs(0) & "');document.getElementById('XGUFRM').submit();return false;"" >"


			XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


			xgurs.movenext
		Loop

		Call ofv.cerrarconsulta(xgurs)


		XGUFBTD1xxx = XGUFBTD1






		Case "TL"

		XGUFILDSC1 = ":&nbsp;LETRA"

		xTrColor = "#F9F9F9"


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If


		Dim ATABXGULTR(28)

		ATABXGULTR(1)  = "A"
		ATABXGULTR(2)  = "B"
		ATABXGULTR(3)  = "C"
		ATABXGULTR(4)  = "D"
		ATABXGULTR(5)  = "E"
		ATABXGULTR(6)  = "F"
		ATABXGULTR(7)  = "G"
		ATABXGULTR(8)  = "H"
		ATABXGULTR(9)  = "I"
		ATABXGULTR(10) = "J"
		ATABXGULTR(11) = "K"
		ATABXGULTR(12) = "L"
		ATABXGULTR(13) = "M"
		ATABXGULTR(14) = "N"
		ATABXGULTR(15) = "O"
		ATABXGULTR(16) = "P"
		ATABXGULTR(17) = "Q"
		ATABXGULTR(18) = "R"
		ATABXGULTR(19) = "S"
		ATABXGULTR(20) = "T"
		ATABXGULTR(21) = "U"
		ATABXGULTR(22) = "V"
		ATABXGULTR(23) = "W"
		ATABXGULTR(24) = "X"
		ATABXGULTR(25) = "Y"
		ATABXGULTR(26) = "Z"
		ATABXGULTR(27) = "NO"
		ATABXGULTR(28) = "NO"


		For XGUINX=1 To UBound(ATABXGULTR) Step 4

			XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If ATABXGULTR(XGUINX) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"

			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
			''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"





			XGUFBTD1 = XGUFBTD1 & "</TD>"


			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If ATABXGULTR(XGUINX+1) = XGUITEMCD Then
				XGUFBTD1 = XGUFBTD1 & XGUSTYLS
			Else
				XGUFBTD1 = XGUFBTD1 & XGUSTYLN
			End If

			XGUFBTD1 = XGUFBTD1 & " >"
			XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
			''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



			XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX+1) & "');document.getElementById('XGUFRM').submit();return false;"" >"





			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If  ATABXGULTR(XGUINX+2) = "NO" Then

				XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

			Else
				If ATABXGULTR(XGUINX+2) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If
				XGUFBTD1 = XGUFBTD1 & " >"
				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
				''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



				XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX+2) & "');document.getElementById('XGUFRM').submit();return false;"" >"



			End If

			XGUFBTD1 = XGUFBTD1 & "</TD>"

			XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

			If  ATABXGULTR(XGUINX+3) = "NO" Then

				XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

			Else
				If ATABXGULTR(XGUINX+3) = XGUITEMCD Then
					XGUFBTD1 = XGUFBTD1 & XGUSTYLS
				Else
					XGUFBTD1 = XGUFBTD1 & XGUSTYLN
				End If
				XGUFBTD1 = XGUFBTD1 & " >"
				XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
				'''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


				XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX+3) & "');document.getElementById('XGUFRM').submit();return false;"" >"




			End If

			XGUFBTD1 = XGUFBTD1 & "</TD>"



			XGUFBTD1 = XGUFBTD1 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If

		Next



		Case "TX"

		XGUFILDSC1 = ":&nbsp;TEXTO"



		xTrColor = "#FFFFFF"


		If XGUCODIGO <> "NO" Then
			XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
		XGUFBTD1 = XGUFBTD1 & " >"

		XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

		If XGUITEMCD <> "NO" Then
			XGUFBTD1 = XGUFBTD1 & XGUSTYLS
		Else
			XGUFBTD1 = XGUFBTD1 & XGUSTYLN
		End If

		XGUFBTD1 = XGUFBTD1 & " >"
		XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE='" & XGUITEMCD & "' CLASS=DT size=40 >&nbsp;&nbsp;"

		XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
		''    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


		XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TX','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO',document.getElementById('XGUBTEXTO').value);document.getElementById('XGUFRM').submit();return false;"" >"




		XGUFBTD1 = XGUFBTD1 & "</TD>"

		XGUFBTD1 = XGUFBTD1 & "</TR>"





		Case Else

		XGUFILDSC1 = "&nbsp;"

		XGUFILTRO = "NO"
		XGUITEMCD = "NO"
		XGUCODIGO = "NO"
		XGUCODDSC = "NO"


	End Select






	XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
	XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTD1 & "</TABLE>"


	XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
	XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

	XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
	XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
	XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

	XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
	XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
	XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

	XGUCONTENT = ""
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
	XGUCONTENT = XGUCONTENT & "<td class=TTR width=50% >&nbsp;COMPONENTES</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"

	XGUCONTENTM = XGUFBDIV1

	' ------------------------


	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' ARMADO DE DIV SEGUN NIVEL DE PROFUNDIDAD ...
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Select Case XGUMASTER


		'  IF XGUCODIGO <> "NO" THEN
		'     XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
		'  ELSE
		'     XGUFNC2 = ""
		'  END IF

		Case "TI","TL","TX"




		Call GETLISTCMP(XGUMASTER,"0",XGUITEMCD)


		XGUTABKEY  = Split(XGUTABLISKEY,Chr(9))
		XGUTABKEY2 = Split(XGUTABLISKEY2,Chr(9))
		XGUTABDAT  = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"



		If XGUFCOBJ <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR COMPONENTE.'  ID='---" & XGUTABKEY(XGUINX) & "' "

			If CStr(XGUTABKEY(XGUINX)) = CStr(XGUFCOBJ) Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If



			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next





		Case "0"

		Call GETLISTCMP("L","0","0")
		XGUTABKEY  = Split(XGUTABLISKEY,Chr(9))
		XGUTABKEY2 = Split(XGUTABLISKEY2,Chr(9))
		XGUTABDAT  = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"



		If XGUFCOBJ <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR COMPONENTE.'  ID='---" & XGUTABKEY(XGUINX) & "' "

			If CStr(XGUTABKEY(XGUINX)) = CStr(XGUFCOBJ) Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If



			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next




		Case "1"


		Call GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
		XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLPCMP("L",XGUFCOBJ,XGUFTOBJ,"0")
		XGUTABKEY = Split(XGUTABLISKEY,Chr(9))
		XGUTABDAT = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"


		If XGUFNIV1 <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFNIV1 & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR SOLAPA.'  ID='---" & XGUTABKEY(XGUINX) & "' "


			If CStr(XGUTABKEY(XGUINX)) = XGUFNIV1 Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If

			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUTABKEY(XGUINX) & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next



		Case "2"

		Call GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
		XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF1 = XGUFBTDF1 & "<TR  "
		XGUFBTDF1 = XGUFBTDF1 & " >"

		XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
		XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
		XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF1 = XGUFBTDF1 & "</TR>"

		Call GETLISTNV2CMP("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,"0")
		XGUTABKEY = Split(XGUTABLISKEY,Chr(9))
		XGUTABDAT = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"


		If XGUFNIV2 <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFNIV2 & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If


		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 1ER. NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "


			If CStr(XGUTABKEY(XGUINX)) = XGUFNIV2 Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If

			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUTABKEY(XGUINX) & "','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next





		Case "3"

		Call GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
		XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF1 = XGUFBTDF1 & "<TR  "
		XGUFBTDF1 = XGUFBTDF1 & " >"

		XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
		XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
		XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF1 = XGUFBTDF1 & "</TR>"



		Call GETLISTNV2CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF2 = XGUFBTDF2 & "<TR  "
		XGUFBTDF2 = XGUFBTDF2 & " >"

		XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
		XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
		XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF2 = XGUFBTDF2 & "</TR>"





		Call GETLISTNV3CMP("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,"0")
		XGUTABKEY = Split(XGUTABLISKEY,Chr(9))
		XGUTABDAT = Split(XGUTABLISDAT,Chr(9))

		xTrColor = "#FFFFFF"

		If XGUFNIV3 <> "0" Then
			XGUFNC2 = " document.getElementById('---" & XGUFNIV3 & "').scrollIntoView(true); " & Chr(13)
		Else
			XGUFNC2 = ""
		End If



		For XGUINX=0 To UBound(XGUTABKEY)


			XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
			XGUFBTD2 = XGUFBTD2 & " >"

			XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 2DO.NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



			If CStr(XGUTABKEY(XGUINX)) = XGUFNIV3 Then
				XGUFBTD2 = XGUFBTD2 & XGUSTYLS
			Else
				XGUFBTD2 = XGUFBTD2 & XGUSTYLN
			End If


			XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('4','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUTABKEY(XGUINX) & "','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
			XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
			XGUFBTD2 = XGUFBTD2 & "</TR>"


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If


		Next


		Case "4"

		Call GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF0 = XGUFBTDF0 & "<TR  "
		XGUFBTDF0 = XGUFBTDF0 & " >"

		XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
		XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
		XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF0 = XGUFBTDF0 & "</TR>"


		Call GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF1 = XGUFBTDF1 & "<TR  "
		XGUFBTDF1 = XGUFBTDF1 & " >"

		XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
		XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
		XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF1 = XGUFBTDF1 & "</TR>"



		Call GETLISTNV2CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF2 = XGUFBTDF2 & "<TR  "
		XGUFBTDF2 = XGUFBTDF2 & " >"

		XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
		XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
		XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF2 = XGUFBTDF2 & "</TR>"



		Call GETLISTNV3CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,XGUFNIV3)
		XGUFCOBJKEY = XGUTABLISKEY
		XGUFCOBJDAT = XGUTABLISDAT


		XGUFBTDF3 = XGUFBTDF3 & "<TR  "
		XGUFBTDF3 = XGUFBTDF3 & " >"

		XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT  "
		XGUFBTDF3 = XGUFBTDF3 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUFNIV3 & "','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
		XGUFBTDF3 = XGUFBTDF3 & XGUImG & "&nbsp;"
		XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
		XGUFBTDF3 = XGUFBTDF3 & "</TR>"





		XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
		XGUFBTD2 = XGUFBTD2 & " >"
		XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI "
		XGUFBTD2 = XGUFBTD2 & XGUSTYLN
		XGUFBTD2 = XGUFBTD2 & "  >"
		XGUFBTD2 = XGUFBTD2 & "----------------------------------------------------------------------</TD>"
		XGUFBTD2 = XGUFBTD2 & "</TR>"


	End Select



	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



	'XGUMASTER

	If XGUMASTER > "0" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


	End If

	If XGUMASTER > "1" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


	End If

	If XGUMASTER > "2" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


	End If

	If XGUMASTER > "3" Then

		XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
		XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


	End If

	If XGUMASTER > "0" Then

		XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;width:99%;background-color:#ffffff;"
		XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;'>"
		XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

	End If




	XGUFBTAB2 = XGUFBTAB2 & "<table id=""MTBSUSR5""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
	XGUFBTAB2 = XGUFBTAB2 & " >" & XGUFBTD2 & "</TABLE>"

	XGUFBDIV2  = "<DIV id='BSUSR5' STYLE='margin-top:1;margin-left:1;height:220;width:99%;background-color:#ffffff;"
	XGUFBDIV2  =  XGUFBDIV2  & "border-style:solid;border-width:1;border-color:#eeeeee;overflow:auto;'>"
	XGUFBDIV2  =  XGUFBDIV2  & XGUFBTAB2 & "</DIV>"

	XGUCONTENT = ""
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV1 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=TTR width=100% >&nbsp;SELECCION" & XGUFILDSC1 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"
	XGUCONTENT = XGUCONTENT & "<tr>"
	XGUCONTENT = XGUCONTENT & "<td class=ut   >" & XGUFBDIV2 & "</td>"
	XGUCONTENT = XGUCONTENT & "</tr>"


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	Call ofv.cerrarconn(xgucn)


	Whtml = ""


	Whtml = Whtml & "<center>"
	Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:650;"
	Whtml = Whtml & "'>"
	Whtml = Whtml & "<CENTER>"




	Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=99% cellspacing=0 cellpadding=2 >"
	Whtml = Whtml & "<tr HEIGHT=25 >"
	Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;&nbsp;SELECCIONAR&nbsp;COMPONENTES&nbsp;Y&nbsp;CONTENIDO&nbsp;</td>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	Whtml = Whtml & "<table id=""MTBSUSR9""  class='aaa' width=99%  cellspacing=0 cellpadding=0 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"
	Whtml = Whtml & "<tD ALIGN=RIGHT >"
	Whtml = Whtml & XGUCONTEMS
	Whtml = Whtml & "</tD>"
	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"


	Whtml = Whtml & "<table id=""MTBSUSR2""  class='aaa' width=99%  cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"
	Whtml = Whtml & "<tr>"

	If Session("EMPMUN")  = "S" Then
		' XGUFBBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESA : '" & XGUCODDSC  &  "  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','TE','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
		'  XGUFBBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI5 & "</FORM>"
		Whtml = Whtml & XGUFBBOT5
	End If
	Whtml = Whtml & XGUFBBOT1
	Whtml = Whtml & XGUFBBOT2
	Whtml = Whtml & XGUFBBOT3
	Whtml = Whtml & XGUFBBOT4
	'Whtml = Whtml & XGUFBBOT5

	If XGUCODDSC <> "NO" Then
		Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
	Else
		Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
	End If

	Whtml = Whtml & "</tr>"
	Whtml = Whtml & "</table>"

	''XGUCONTENTM

	'Whtml = Whtml & "<table id=""MTBSUSR3B""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
	'  Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	'  Whtml = Whtml & " >"
	'
	'  Whtml = Whtml & XGUFBTD1xxx
	'
	'  Whtml = Whtml & "</table>"





	Whtml = Whtml & "<table id=""MTBSUSR3""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
	Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
	Whtml = Whtml & " >"

	Whtml = Whtml & XGUCONTENT

	Whtml = Whtml & "</table>"
	Whtml = Whtml & "</CENTER>"
	Whtml = Whtml & "</DIV>"
	Whtml = Whtml & "</center>"



	Whtml = Whtml & "<FORM NAME='XGUFRM' id='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' id='XGUMASTER' >"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFCOBJ' id='XGUFCOBJ'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFTOBJ' id='XGUFTOBJ'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV1' id='XGUFNIV1'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV2' id='XGUFNIV2'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV3' id='XGUFNIV3'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV4' id='XGUFNIV4'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFNIV5' id='XGUFNIV5'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFILTRO' id='XGUFILTRO'>"
	Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUITEMCD' id='XGUITEMCD'>"

	Whtml = Whtml & "</FORM>"

	Whtml = Whtml & "<script language=javascript >"
	Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4,XGU5,XGU6,XGU7,XGU8,XGU9) { " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFCOBJ').value=XGU1; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFTOBJ').value=XGU2; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV1').value=XGU3; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV2').value=XGU4; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV3').value=XGU5; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV4').value=XGU6; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFNIV5').value=XGU7; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU8; " & Chr(13)
	Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU9; " & Chr(13)

	Whtml = Whtml & " } " & Chr(13)
	Whtml = Whtml & XGUFNC1
	Whtml = Whtml & XGUFNC2
	'whtml = Whtml & "  alert('12');"
	Whtml = Whtml & "</script>"

	GenBuscaCMP = Whtml


End Function






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function SETEMPRE()

	xGRP = "S"
	xUSR = "S"

	Set xcn = ofv.conectar(ofv.strconn2)

	Strsql = "select SITE from USUARIOS where ID = '" & SESSION("USUARIO") & "'"
	SESSION("ZZZ")=STRSQL
	Set Xrs = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
	If Not Xrs.eof Then
		USREMP = Xrs(0)
		USREMP = CLng(USREMP)
	Else
		USREMP=session("SGU_site")
	End If

	'Strsql = "select * from empreatrm where ATRMTIPO = 'G' AND atrmusrgrp = '" & SESSION("SEGGRP") & "' ORDER BY EMPRE "}
	xgrupos=Split(session("SegGrpLst"),Chr(9))
	sgrupo=""
	delim=""
	For i =0 To UBound(xgrupos)
		sgrupo=sgrupo & delim & "'" & xgrupos(i) & "'"
		delim=","
	Next
	If sgrupo="" Then
		sgrupo= "'" & Session("SegGrp") & "'"
		i=1
	End If
	' CALCULO EL MAX atrmver Y EL TOTAL DE EMPRESAS
	' SI EL TOTAL DE EMPRESA ES IGUAL LA CANTIDAD DE GRUPOS TOMO MAX ATRMVER

	Strsql = " select  max(empre),max( cast(atrmver as int)) as atrmver ,max( cast(atrmact as int)),count(empre) as iempre  from empreatrm where ATRMTIPO = 'G' AND atrmusrgrp in (" & sgrupo & ")  "

	Strsql = Strsql & " group by empre "
	'response.write  Session("SegGrp")
	'SESSION("ZZZ")=STRSQL
	Set Xrs1 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
	If Xrs1.eof Then
		xGRP = "N"
	End If


	Strsql = "select * from empreatrm where ATRMTIPO = 'U' AND atrmusrgrp = '" & SESSION("USUARIO") & "' and atrmver <> 0 ORDER BY EMPRE "
	Set Xrs2 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
	If Xrs2.eof Then
		xUSR = "N"
	End If
	'response.write Strsql
	ATREMP(0,0) = 0
	'IAEMP=0
	'SESSION("ZZZ")= xGRP & "   " &    xUSR & "   " & i
	Strsql = "select * from CLIENTEPLAN ORDER BY ID"
	Set Xrs = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
	If Not Xrs.eof Then
		Do While Not XRS.EOF
			IAEMP = IAEMP + 1
			TAEMP = IAEMP

			ATREMP(IAEMP,0) = XRS(0)
			ATREMP(IAEMP,1) = XRS(6)
			ATREMP(IAEMP,4) = XRS(1)
			ATREMP(IAEMP,5) = XRS(5)
			ATREMP(IAEMP,2) = 1
			ATREMP(IAEMP,3) = 1

			For XX1=0 To 50
				If XGRP <> "N" Then
					If  XRS1(0) = XRS(0) Then
						' CANT DE GRUPOS USR IGUAL A CANT DE EMPRESAS
						If  CInt(i)  = CInt(XRS1("iempre")) Then
							ATREMP(IAEMP,2) = Abs(CInt(XRS1(1)))
							'   end if
							ATREMP(IAEMP,3) = Abs(CInt(XRS1(2)))
						End If
						XRS1.MOVENEXT
						If XRS1.EOF Then XGRP = "N"
						Exit For
					ElseIf  XRS1(0) < XRS(0) Then
						XRS1.MOVENEXT
						If XRS1.EOF Then XGRP = "N"
					End If
				Else
					Exit For
				End If
			Next

			For XX1=0 To 50
				If XUSR <> "N" Then
					If  XRS2(0) = XRS(0) Then
						'  If  CInt(i)  = CInt(XRS1("iempre")) Then
						ATREMP(IAEMP,2) = Abs(CInt(XRS2(5)))
						ATREMP(IAEMP,3) = Abs(CInt(XRS2(6)))
						'	End If
						XRS2.MOVENEXT
						If XRS2.EOF Then XUSR = "N"
						Exit For
					ElseIf  XRS2(0) < XRS(0) Then
						XRS2.MOVENEXT
						If XRS2.EOF Then XUSR = "N"
					End If
				Else
					Exit For
				End If
			Next




			If XRS(0) = USREMP Then
				ATREMP(IAEMP,2) = 1

				ATREMP(0,0) = ATREMP(IAEMP,0)
				ATREMP(0,1) = ATREMP(IAEMP,1)
				ATREMP(0,4) = ATREMP(IAEMP,4)
				ATREMP(0,5) = ATREMP(IAEMP,5)
				ATREMP(0,2) = ATREMP(IAEMP,2)
				ATREMP(0,3) = ATREMP(IAEMP,3)


			End If


			XRS.MOVENEXT
		Loop

		If ATREMP(0,0) = 0 Then

			For IAEMP = 1 To TAEMP
				If  ATREMP(IAEMP,2) = 1 Then

					ATREMP(0,0) = ATREMP(IAEMP,0)
					ATREMP(0,1) = ATREMP(IAEMP,1)
					ATREMP(0,4) = ATREMP(IAEMP,4)
					ATREMP(0,5) = ATREMP(IAEMP,5)
					ATREMP(0,2) = ATREMP(IAEMP,2)
					ATREMP(0,3) = ATREMP(IAEMP,3)
					Exit For
				End If


			Next
		End If

		Dim WATREMPEL(5)
		Dim WATREMPTB(50)
		Dim empreok(50)

		For IAEMP = 0 To TAEMP
			empreok(IAEMP)="0"
		Next
		session("EmpHbl")=""
		delim=""
		For IAEMP = 0 To TAEMP
			For IAEMP2 = 0 To 5
				WATREMPEL(IAEMP2) = CStr(ATREMP(IAEMP,IAEMP2) &"")
				If  IAEMP2 = 0 And ATREMP(IAEMP,2)=1  And  Not EstaEmpre(empreok,TAEMP,CStr(ATREMP(IAEMP,0) &"")) Then
					session("EmpHbl")=session("EmpHbl") & delim &  CStr(ATREMP(IAEMP,IAEMP2) &"")
					empreok(IAEMP)=CStr(ATREMP(IAEMP,IAEMP2) &"")
					delim=Chr(9)
				End If
			Next

			WATREMPTB(IAEMP) = Join(WATREMPEL,Chr(09))

		Next



		SESSION("ATREMP") = Join(WATREMPTB,Chr(10))
		SESSION("TAEMP")  = TAEMP
		SESSION("SETEMP") = "S"


	End If



End Function


Function EstaEmpre(a,t,e)
	EstaEmpre=False
	For xxx = 0 To t
		If CStr(a(xxx)& "") =CStr(e) Then
			EstaEmpre=True
			Exit For
		End If
	Next
End Function


Function NEWEMPRE(XEMPRE)

	If XEMPRE = "NO" Then

		If SESSION("SETEMP") <> "S" Then   xxxx = SETEMPRE()

		WATREMPTB = Split(SESSION("ATREMP"),Chr(10))

		For IAEMP=0 To UBound(WATREMPTB)

			WATREMPEL = Split(WATREMPTB(IAEMP),Chr(09))

			For IAEMP2=0 To UBound(WATREMPEL)

				ATREMP(IAEMP,IAEMP2) = WATREMPEL(IAEMP2)

			Next

		Next

		TAEMP  = SESSION("TAEMP")

		NEWEMPRE = "OK"

	Else

		NEWEMPRE = "N"

		If ATREMP(0,0) <> XEMPRE Then

			NEWEMPRE = "S"

			For IAEMP = 1 To TAEMP
				If  ATREMP(IAEMP,0) = XEMPRE Then

					ATREMP(0,0) = ATREMP(IAEMP,0)
					ATREMP(0,1) = ATREMP(IAEMP,1)
					ATREMP(0,4) = ATREMP(IAEMP,4)
					ATREMP(0,5) = ATREMP(IAEMP,5)
					ATREMP(0,2) = ATREMP(IAEMP,2)
					ATREMP(0,3) = ATREMP(IAEMP,3)
					Exit For
				End If


			Next

			Dim WATREMPEL(5)
			Dim WATREMPTB(50)


			For IAEMP = 0 To TAEMP
				For IAEMP2 = 0 To 5
					WATREMPEL(IAEMP2) = ATREMP(IAEMP,IAEMP2)
				Next

				WATREMPTB(IAEMP) = Join(WATREMPEL,Chr(09))

			Next


			SESSION("ATREMP") = Join(WATREMPTB,Chr(10))
			SESSION("TAEMP")  = TAEMP

		End If
	End If

End Function


Function GETGLBHMSW(XITIME)

	If Not IsNumeric(XITIME & "") Then

		GETGLBHMSW = "000005"

	Else

		XITIME=CLng("0" & Replace(CStr(XITIME & ""),"-",""))
		GLBbb = XITIME Mod 3600
		GLBbC = (GLBBB) Mod 60
		GLBbD = (GLBBC) Mod 1

		GLBSH = (XITIME - GLBBB) / 3600
		GLBSM = (GLBBB - GLBBC) / 60
		GLBSS = GLBBC / 1

		GLBSH = CStr(GLBSH)
		GLBSM = CStr(GLBSM)
		GLBSS = CStr(GLBSS)

		If Len(GLBSH) = 1 Then GLBSH = "0" & GLBSH
		If Len(GLBSM) = 1 Then GLBSM = "0" & GLBSM
		If Len(GLBSS) = 1 Then GLBSS = "0" & GLBSS

		XOSH = GLBSH & ":"
		XOSM = GLBSM & ":"
		XOSS = GLBSS & ""


		GETGLBHMSW = XOSH & XOSM & XOSS

	End If

End Function


Function GETGLBHMS(XITIME)

	If Not IsNumeric(XITIME & "") Then

		GETGLBHMS = "000005"

	Else
		''response.write "-" & XITIME & "-" & "<br>"
		XITIME=CLng("0" & Replace(CStr(XITIME & ""),"-",""))
		GLBbb = XITIME Mod 3600
		GLBbC = (GLBBB) Mod 60
		GLBbD = (GLBBC) Mod 1

		GLBSH = (XITIME - GLBBB) / 3600
		GLBSM = (GLBBB - GLBBC) / 60
		GLBSS = GLBBC / 1

		If GLBSH > 0 Then XOSH = GLBSH & "h&nbsp;"
		If GLBSM > 0 Then XOSM = GLBSM & "m&nbsp;"
		If GLBSS > 0 Then XOSS = GLBSS & "s&nbsp;"


		GETGLBHMS = XOSH & XOSM & XOSS

	End If

End Function

Function GETFECH()



	sh = DatePart("h",Time())
	sm = DatePart("n",Time())
	ss = DatePart("s",Time())

	XLHORDAT = (CInt(sh)*3600) + (CInt(sm)*60) + CInt(ss)
	''   XLHORDAT = TIME()
	''  RESPONSE.WRITE XLHORDAT & "<BR>"
	XLFECDAT = CLng(Date())



End Function

Function GETNWEN(XNCODIGO,XNTIPO)


	sh = DatePart("h",Time())
	sm = DatePart("n",Time())
	ss = DatePart("s",Time())

	XCodobJ = sh & sm & ss

	GThms   = DateDiff("d",Date(),CDate("12/12/2050")) & FormatDateTime(Now(),3)
	GTses   = session.sessionid
	GTsecu  = CStr(XCodobJ & GThms & GTses)

	GTsecuf = XNCODIGO & XNTIPO


	For GTi=1 To Len(GTsecu)
		GTcs = Mid(GTsecu,GTi,1)
		If GTcs => "0" And GTcs <= "9" Then
			GTsecuf = GTsecuf & GTcs
		End If
	Next

	GETNWEN = Trim(GTsecuf)

End Function

Function GETLOGN(XLTYPE)

	If XLTYPE = "" Then XLTYPE = "L"

	sh = DatePart("h",Time())
	sm = DatePart("n",Time())
	ss = DatePart("s",Time())

	XCodobJ = sh & sm & ss

	GThms = DateDiff("d",Date(),CDate("12/12/2050")) & FormatDateTime(Now(),3)
	GTses = session.sessionid
	GTsecu = CStr(XCodobJ & GThms & GTses)

	Select Case XLTYPE
		Case "L"
		GTsecuf = "LOG"
		Case "LE"
		GTsecuf = "LXT"
		Case "LD"
		GTsecuf = "LDE"
	End Select

	For GTi=1 To Len(GTsecu)
		GTcs = Mid(GTsecu,GTi,1)
		If GTcs => "0" And GTcs <= "9" Then
			GTsecuf = GTsecuf & GTcs
		End If
	Next

	GETLOGN = Trim(GTsecuf)

End Function


'''' LOG DE USUARIOS ''''''''''''''''''''''''''''''''''''''''''''''''''

Function GENUSRLOG()

	'      Session("Usuario") = UCase(Request.Form("Usuarnt"))
	'      Session("UsuarNT") = UCase(Request.Form("Usuarnt"))
	'      Session("DOMINIO") = UCase(Request.Form("DOMAINnt"))
	'      Session("TRUSTEDUSR") = "S"


	StrSql = "Select * From OBJALOG where CODOBJETO = 0 AND  TIPOBJETO = 'SKRLOG' "
	STRSQL = STRSQL & " AND USUARIO = '" & session("USUARIO") & "' AND  NETUSUARIO = '" & session("USUARNT") & "' "
	STRSQL = STRSQL & " AND DOMINIO = '" & session("DOMINIO") & "' "
	''       response.write strsql & "<BR>"
	Set Xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	If Xrs.eof Then

		'if session("APLMDCLI")="BC" THEN
		STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED,USRNAME,CODEMP,CODSECTOR) VALUES ("
		STRSQL = STRSQL & "'0','SKRLOG','NO','0','0','0','0','" & session("USUARIO") & "','" & session("USUARNT") & "','" & session("DOMINIO") & "','" & Session("TRUSTEDUSR") & "','"&Session("SegUsuario")&"','" & session("CLIEMP") & "','" & Session("UsrSector") & "')"
		'ELSE
		'STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED,CODEMP) VALUES ("
		'STRSQL = STRSQL & "'0','SKRLOG','NO','0','0','0','0','" & session("USUARIO") & "','" & session("USUARNT") & "','" & session("DOMINIO") & "','" & Session("TRUSTEDUSR") & "','" & session("CLIEMP") & "')"
		'END IF
		Call ofv.crearconsultaEx(StrSql,cn2,1,parametros)
		'        response.write strsql

	End If



End Function

Function GETUSRLOGUPD(XLTIPO)

	'' RESPONSE.WRITE XLTIPO & "<BR>"
	XLOPC = "N"

	Select Case XLTIPO

		Case "LOGON"
		''  RESPONSE.WRITE XLTIPO & "0<BR>"
		XLOPC = "S"

		XLLOGIDDET = GETLOGN("L")
		SESSION("LOGIDUSRDET") = XLLOGIDDET
		XXXX = GETFECH()
		XXXX = GENUSRLOGDET(XLTIPO)
		XLUPD = " SET estadou=0,mail='" & Session("UsrMail") &"' ,CODEMP='" & session("CLIEMP") & "', AEDI = AEDI + 1, FULTACC = '" & XLFECDAT & "' "
		Sgrpuser=""
		If session("SegGrpLst")= "" Then
			XLUPD = XLUPD  & ",grupos='" & Session("SegGrp") & "'"
		Else
			XLUPD = XLUPD  & ",grupos='"  & session("SegGrpLst") & "'"
		End  If

		'XLUPD = XLUPD  & ",codsector='" & Session("UsrSector") & "'"
		XLUPD = XLUPD  & ",codpuesto='" & Session("UsrPuesto") & "'"

		Case "LOGOFF"
		XLOPC = "S"
		XXXX = GETFECH()
		XLHORAINI = GETUSRLOGDETTIME()
		XLUPD = "  SET estadou=0,mail='" & Session("UsrMail") &"' , CODEMP='" & session("CLIEMP") & "', TTOTALACC = TTOTALACC + " & (CDbl("0" & XLHORDAT) - CDbl("0" & XLHORAINI))
		If session("SegGrpLst")= "" Then
			XLUPD = XLUPD  & ",grupos='" & Session("SegGrp")  & "'"
		Else
			XLUPD = XLUPD  & ",grupos='" & session("SegGrpLst")  & "'"
		End  If
		' XLUPD = XLUPD  & ",codsector='" & Session("UsrSector") & "'"
		XLUPD = XLUPD  & ",codpuesto='" & Session("UsrPuesto") & "'"

		XXXX = GENUSRLOGDET(XLTIPO)


	End Select
	'' RESPONSE.WRITE SESSION("LOGIDUSRDET") & "1<BR>"
	If XLOPC = "S" Then
		'' RESPONSE.WRITE SESSION("LOGIDUSRDET") & "2<BR>"
		'if session("APLMDCLI")="BC" THEN
		StrSql = "UPDATE OBJALOG " & XLUPD &",usrname='" & Session("SegUsuario") & "', CODSECTOR='" & Session("UsrSector") & "' "
		'ELSE
		'StrSql = "UPDATE OBJALOG " & XLUPD
		'END IF

		StrSql = StrSql & " where CODOBJETO = 0 AND  TIPOBJETO = 'SKRLOG' "
		STRSQL = STRSQL & " AND USUARIO = '" & session("USUARIO") & "' AND  NETUSUARIO = '" & session("USUARNT") & "' "
		STRSQL = STRSQL & " AND DOMINIO = '" & session("DOMINIO") & "' "
		' response.write strsql & "<BR>" & Session("SegGrp")
		Call ofv.crearconsultaEx(StrSql,cn2,1,parametros)

		'response.write StrSql

	End If



End Function

Function GENUSRLOGDET(XLTIPO)

	XLOPC = "N"

	XLDTACC = "NO"
	XLDDESD = Session("Form")
	XLDFINI = "NO"
	XLDHINI = "0"
	XLDFFIN = "NO"
	XLDHFIN = "0"
	XLDUSUA = SESSION("USUARIO")
	XLDNUSU = Session("UsuarNT")
	XLDGRUP = SESSION("SEGGRP")
	XLDSESS = SESSION.SESSIONID
	XLDDIRI = UCase(Request.ServerVariables("REMOTE_ADDR"))
	XLDLOGI = SESSION("LOGIDUSRDET")
	XLDDOMI = Session("DOMINIO")
	XLDTRUS = Session("TRUSTEDUSR")

	Select Case XLTIPO

		Case "LOGON"
		XLOPC = "S"
		XLDTACC = "LOGON"
		XLDFINI = XLFECDAT
		XLDHINI = XLHORDAT
		XLDFFIN = XLFECDAT
		XLDHFIN = XLHORDAT


		Case "LOGOFF"



		'   StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
		'  StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
		''  response.write strsql   & "<BR>"
		'   CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)





	End Select



	If XLOPC = "S" Then



		STRSQL = "INSERT INTO OBJALOGDET ("
		STRSQL = STRSQL & "CODOBJETO,TIPOBJETO,TACC,DESDEOBJ,FINICIO,HINICIO,FFINAL,"
		STRSQL = STRSQL & "HFINAL,USUARIO,NETUSUARIO,GRUPO,SESSIONID,DIRIP,LOGID,DOMINIO,TRUSTED,USRLOGID) VALUES ("
		STRSQL = STRSQL & "'0','SKRLOG',"
		STRSQL = STRSQL & "'" & XLDTACC & "','" & XLDDESD & "',"
		STRSQL = STRSQL & "'" & XLDFINI & "','" & XLDHINI & "',"
		STRSQL = STRSQL & "'" & XLDFFIN & "','" & XLDHFIN & "',"
		STRSQL = STRSQL & "'" & XLDUSUA & "','" & XLDNUSU & "',"
		STRSQL = STRSQL & "'" & XLDGRUP & "','" & XLDSESS & "',"
		STRSQL = STRSQL & "'" & XLDDIRI & "','" & XLDLOGI & "','" & XLDDOMI & "','" & XLDTRUS & "','" & XLDLOGI & "')"
		''      response.write strsql   & "<BR>"
		Call ofv.crearconsultaEx(StrSql,cn2,1,parametros)




	End If


	StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
	StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
	''  response.write strsql   & "<BR>"
	Call ofv.crearconsultaEx(StrSql,cn2,1,parametros)



End Function

Function GETUSRLOGDETTIME()

	GETUSRLOGDETTIME = XLHORDAT




	StrSql = "Select HINICIO From OBJALOGDET where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
	'   response.write strsql   & "<BR>"
	Set Xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	If Not Xrs.eof Then

		GETUSRLOGDETTIME = XRS(0)

	End If



End Function

'''' LOG DE OBJETOS ''''''''''''''''''''''''''''''''''''''''''''''''''


Function GENOBJALOG()

	StrSql = "Select * From OBJALOG where CODOBJETO = " & session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
	''       response.write strsql & "<BR>"
	Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Xrs.eof Then

		STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED) VALUES ("
		STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TCliente") & "','NO','0','0','0','0','NO','NO','NO','NO')"

		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)
		'        response.write strsql

	End If



End Function

Function GETOBJALOGUPD(XLTIPO)

	'' RESPONSE.WRITE XLTIPO & "<BR>"
	XLOPC = "N"

	Select Case XLTIPO

		Case "ED"
		''  RESPONSE.WRITE XLTIPO & "0<BR>"
		XLOPC = "S"

		XLLOGIDDET = GETLOGN("LD")
		SESSION("LOGIDDET") = XLLOGIDDET
		XXXX = GETFECH()
		XXXX = GENLOGDET(XLTIPO)
		XLUPD = " SET AEDI = AEDI + 1, FULTACC = '" & XLFECDAT & "' "

		Case "CON","ACCMOD"


		If XLTIPO = "CON" Then XLOPC = "S"

		XLLOGIDDET = GETLOGN("LD")
		SESSION("LOGIDDET") = XLLOGIDDET
		XXXX = GETFECH()
		XXXX = GENLOGDET(XLTIPO)
		XLUPD = " SET ACON = ACON + 1, FULTACC = '" & XLFECDAT & "' "

		Case "ACT"
		XLOPC = "S"
		XLUPD = " SET CACT = CACT + 1 "

		Case "FIN"
		XLOPC = "S"
		XXXX = GETFECH()
		XLHORAINI = GETLOGDETTIME()
		XLUPD = " SET TTOTALACC = TTOTALACC + " & (CLng(XLHORDAT) - CLng(XLHORAINI))
		XXXX = GENLOGDET(XLTIPO)


	End Select
	'' RESPONSE.WRITE SESSION("LOGIDDET") & "1<BR>"
	If XLOPC = "S" Then
		'' RESPONSE.WRITE SESSION("LOGIDDET") & "2<BR>"

		StrSql = "UPDATE OBJALOG " & XLUPD
		StrSql = StrSql & " where CODOBJETO = " & session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
		''    response.write strsql & "<BR>"
		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)



	End If



End Function



Function GENLOGDET(XLTIPO)

	XLOPC = "N"

	XLDTACC = "NO"
	XLDDESD = Session("Form")
	XLDFINI = "NO"
	XLDHINI = "0"
	XLDFFIN = "NO"
	XLDHFIN = "0"
	XLDUSUA = SESSION("USUARIO")
	XLDNUSU = Session("UsuarNT")
	XLDGRUP = SESSION("SEGGRP")
	XLDSESS = SESSION.SESSIONID
	XLDDIRI = UCase(Request.ServerVariables("REMOTE_ADDR"))
	XLDLOGI = SESSION("LOGIDDET")
	XLDDOMI = Session("DOMINIO")
	XLDTRUS = Session("TRUSTEDUSR")
	XLDULOG = SESSION("LOGIDUSRDET")

	Select Case XLTIPO

		Case "ED"
		XLOPC = "S"
		XLDTACC = "EDEN"
		XLDFINI = XLFECDAT
		XLDHINI = XLHORDAT
		XLDFFIN = XLFECDAT
		XLDHFIN = XLHORDAT

		Case "CON"
		XLOPC = "S"
		XLDTACC = "CONSEN"
		XLDFINI = XLFECDAT
		XLDHINI = XLHORDAT

		Case "ACCMOD"
		XLOPC = "S"
		XLDTACC = "ACCMOD"
		XLDFINI = XLFECDAT2
		XLDHINI = XLHORDAT2
		XLDFFIN = XLFECDAT
		XLDHFIN = XLHORDAT

		Case "FIN"



		'         StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
		'         StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDDET") & "' "
		''  response.write strsql   & "<BR>"
		'         CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)





	End Select



	If XLOPC = "S" Then



		STRSQL = "INSERT INTO OBJALOGDET ("
		STRSQL = STRSQL & "CODOBJETO,TIPOBJETO,TACC,DESDEOBJ,FINICIO,HINICIO,FFINAL,"
		STRSQL = STRSQL & "HFINAL,USUARIO,NETUSUARIO,GRUPO,SESSIONID,DIRIP,LOGID,DOMINIO,TRUSTED,USRLOGID) VALUES ("
		STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TCliente") & "',"
		STRSQL = STRSQL & "'" & XLDTACC & "','" & XLDDESD & "',"
		STRSQL = STRSQL & "'" & XLDFINI & "','" & XLDHINI & "',"
		STRSQL = STRSQL & "'" & XLDFFIN & "','" & XLDHFIN & "',"
		STRSQL = STRSQL & "'" & XLDUSUA & "','" & XLDNUSU & "',"
		STRSQL = STRSQL & "'" & XLDGRUP & "','" & XLDSESS & "',"
		STRSQL = STRSQL & "'" & XLDDIRI & "','" & XLDLOGI & "','" & XLDDOMI & "','" & XLDTRUS & "','" & XLDULOG & "')"
		''      response.write strsql   & "<BR>"
		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)




	End If


	StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
	StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDDET") & "' "
	''  response.write strsql   & "<BR>"
	Call ofv.crearconsultaEx(StrSql,cn,1,parametros)


End Function

Function GETLOGDETTIME()

	GETLOGDETTIME = XLHORDAT




	StrSql = "Select HINICIO From OBJALOGDET where LOGID = '" & SESSION("LOGIDDET") & "' "
	'   response.write strsql   & "<BR>"
	Set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not Xrs.eof Then

		GETLOGDETTIME = XRS(0)

	End If



End Function

Function GENLOGEXT(XLTIPO)

	XLOPC = "N"

	XLDTACC = "NO"
	XXXX    = GETFECH()
	XLDFINI = XLFECDAT
	XLDHINI = XLHORDAT
	XLDLOGE = GETLOGN("LE")
	XLDLOGI = SESSION("LOGIDDET")
	XLDULOG = SESSION("LOGIDUSRDET")

	Select Case XLTIPO

		Case "EDIDET"
		XLOPC   = "S"
		XLDTACC = XLTIPO
		XLDMOBJ = Session("Form2")
		XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
		XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

		If (XLDCOBJ & "") = "" Then XLOPC = "N"

		'' RESPONSE.WRITE REQUEST.QUERYSTRING & "<BR>"

		Case "ACCDET"
		XLOPC   = "S"
		XLDTACC = XLTIPO
		XLDMOBJ = Session("Form2")
		XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
		XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")


	End Select



	If XLOPC = "S" Then


		STRSQL = "INSERT INTO OBJALOGEXT ("
		STRSQL = STRSQL & "CODOBJETO, TIPOBJETO, "
		STRSQL = STRSQL & "LOGID, TOPER, "
		STRSQL = STRSQL & "TABLA, CLAVE, "
		STRSQL = STRSQL & "FOPER, HOPER, "
		STRSQL = STRSQL & "OBSERV, LOGIDEXT, "
		STRSQL = STRSQL & "CODOBJ, MASTEROBJ, "
		STRSQL = STRSQL & "MODOBJ, USRLOGID  "
		STRSQL = STRSQL & ") VALUES ("
		STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TCliente") & "',"
		STRSQL = STRSQL & "'" & XLDLOGI & "','" & XLDTACC & "',"
		STRSQL = STRSQL & "'NO','NO',"
		STRSQL = STRSQL & "'" & XLDFINI & "','" & XLDHINI & "',"
		STRSQL = STRSQL & "'NO','" & XLDLOGE & "',"
		STRSQL = STRSQL & "'" & XLDCOBJ & "','" & XLDMOBJ & "',"
		STRSQL = STRSQL & "'" & XLDOOBJ & "','" & XLDULOG & "')"

		Set cn = ofv.conectar(ofv.strconn2)
		''  response.write strsql   & "<BR>"
		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Call ofv.cerrarconn(cn)


	End If





End Function

Function SqlExLog(XLTIPO,sql,CNEX1,TABLA,wkeyname,VALOR,TIPOPARM)

	If wkeyname="" Then
		StrsqlZX = "Select * From " & Tabla
		Set rsKEY = ofv.crearconsultaEx(StrSqlZX,CNEX1,1,volver)
		wkeyname = rsKEY(0).name
		Call OFV.CERRARCONSULTA(rsKEY	)
	End If

	STRSQLS= "PRC_EXECSQL '" & Replace(SQL,"'","''") & "'"

	Set VV=ofv.crearconsultaEx(StrSqlS,CNEX1,1,volver)
	If cn.errors.count = 0 Then
		If Not IsNull(VV(0)) Then
			IDINSERT=VV(0)
		Else
			IDINSERT=VALOR
		End If
		xxxr=GENLOGEXTDATA(XLTIPO,TABLA,wkeyname,IDINSERT,TIPOPARM)
	End If
	Call OFV.CERRARCONSULTA(VV)
End Function

Function GENLOGEXTDATA(XLTIPO,TABLA,NOMCOL,IDCOL,TIPOPARAM)

	XLOPC = "N"
	TABLA=UCase(tabla &"")
	NOMCOL=UCase(NOMCOL &"")
	XLDTACC = "NO"
	XXXX = GETFECH()
	XLDFINI = XLFECDAT - CLng(CDate("01-01-1900"))
	XLDHINI = XLHORDAT
	XLDLOGE = GETLOGN("LE")
	XLDLOGI = SESSION("LOGIDDET")
	XLDULOG = SESSION("LOGIDUSRDET")
	TIPOPARAM=UCase(TIPOPARAM&"")
	Select Case XLTIPO

		Case "INSERT"
		XLOPC = "S"
		If TIPOPARAM <> "" Then
			XLDTACC = TIPOPARAM & "-" & XLTIPO
		Else
			XLDTACC=XLTIPO
		End If
		XLDMOBJ = Session("CorApli")
		XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
		XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

		'IF (XLDCOBJ & "") = "" THEN XLOPC = "N"

		'' RESPONSE.WRITE REQUEST.QUERYSTRING & "<BR>"

		Case "UPDATE"
		XLOPC = "S"
		If TIPOPARAM <> "" Then
			XLDTACC = TIPOPARAM & "-" & XLTIPO
		Else
			XLDTACC=XLTIPO
		End If
		XLDMOBJ = Session("CorApli")
		XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
		XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

		Case "DELETE"
		XLOPC = "S"
		If TIPOPARAM <> "" Then
			XLDTACC = TIPOPARAM & "-" & XLTIPO
		Else
			XLDTACC=XLTIPO
		End If
		XLDMOBJ = Session("CorApli")
		XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
		XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

		Case Else
		XLOPC = "S"

		If TIPOPARAM <> "" Then
			XLDTACC = TIPOPARAM & "-" & XLTIPO
		Else
			XLDTACC=XLTIPO
		End If

		XLDMOBJ = Session("CorApli")
		XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
		XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

	End Select
	XLDCOBJ=IDCOL
	If XLOPC = "S" Then

		STRSQLX = "INSERT INTO OBJALOGEXT ("
		STRSQLX = STRSQLX & "CODOBJETO, TIPOBJETO, "
		STRSQLX = STRSQLX & "LOGID, TOPER, "
		STRSQLX = STRSQLX & "TABLA, CLAVE, "
		STRSQLX = STRSQLX & "FOPER, HOPER, "
		STRSQLX = STRSQLX & "OBSERV, LOGIDEXT, "
		STRSQLX = STRSQLX & "CODOBJ, MASTEROBJ, "
		STRSQLX = STRSQLX & "MODOBJ, USRLOGID,VALANT,VALACT  "
		STRSQLX = STRSQLX & ") VALUES ("
		'STRSQLX = STRSQLX & "'" & session("NCliente") & "','" & session("TCliente") & "',"
		STRSQLX = STRSQLX & "'0','0',"
		STRSQLX = STRSQLX & "'" & XLDLOGI & "','" & XLDTACC & "',"
		STRSQLX = STRSQLX & "'"&TABLA & "','"& NOMCOL &"',"
		STRSQLX = STRSQLX & "'" & XLDFINI & "','" & XLDHINI & "',"
		STRSQLX = STRSQLX & "'NO','" & XLDLOGE & "',"
		STRSQLX = STRSQLX & "'" & XLDCOBJ & "','" & XLDMOBJ & "',"
		STRSQLX = STRSQLX & "'" & XLDOOBJ & "','" & XLDULOG
		STRSQLX = STRSQLX & "','" & XVALANT &"','" & XVALACT & "')"

		'  SET cnER = ofv.conectar(ofv.strconn2)

		Call ofv.crearconsultaEx(STRSQLX, cn, 1, parametros)

		' CALL ofv.cerrarconn(cnER)

	End If
	' SESSION("STRSQLX")= STRSQLX   &"---" & XLTIPO & "<BR>"

End Function


Function FechadorCor(wkfecha,param1)


	'if wkfecha & "" <> "" and wkfecha <> "&nbsp;" and wkfecha <> "..." then

	Select Case SESSION("TIPOFECHA")


		Case 0
		If IsDate(wkfecha) Then

			fec = CDate(wkfecha)
			ndia = Day(fec)
			nMes = Month(fec)
			anio=Year(fec)

			If ndia < 10 Then ndia="0" & ndia
			If nMes < 10 Then nMes="0" & nMes

			FechadorCor = ndia &"/" & nmes & "/" & anio
		Else
			FechadorCor=wkfecha
		End If

		Case Else
		FechadorCor=wkfecha

	End Select

End Function


Function  getuserDb()

	Dim oContext
	Dim myUser




	Dirc=  Request.ServerVariables("PATH_TRANSLATED")
	Dirc=Left(Dirc,InStrRev(Dirc,"\",-1))

	Dirc=Dirc & "incfile\Application.xml"

	'Dirc="C:\contextdelivery\Application.xml"

	' response.write dirc
	Set oContext = Server.CreateObject("ContextDeliveryCOM.contextdelivery")
	Set myUser	 =server.CreateObject("ContextDeliveryCOM.User")



	oContext.SetApplicationFile CStr(Dirc)

	Set myUser = OContext.GetUser("AliasAplicacion", "SQLALIAS", "DESA")
	' response.write "Usuario=" & myUser.GetUserName() & "<br>"
	' response.write "Password=" & myUser.GetPassword() & "<br>

	ofv.ADbUid=myUser.GetUserName()

	ofv.ADbPwd =ofv.getCript("E",myUser.GetPassword()&"")

	Set oContext = Nothing
	Set myUser	 = Nothing

End Function



Function GTESTADODSC(GTESTADCOD)

	If CStr(GTESTADCOD & "") = "0" Or CStr(GTESTADCOD & "") = "" Then

		GTESTADODSC  = "-----"

	Else

		Set cnXX = OFV.conectar(OFV.strconn2)

		XXstrsql = "SELECT ESTADO FROM ESTADOS    "
		XXstrsql = XXstrsql & " WHERE OBJID = 0" & GTESTADCOD
		'''   response.write XXstrsql & "<br>"
		Set XXRS = ofv.crearconsultaEx(XXStrSql,cnXX,1,parametros)
		If Not XXRS.eof Then
			XXWKESTADO = XXRS(0)
		Else
			XXWKESTADO = GTESTADCOD
		End If

		GTESTADODSC  = XXWKESTADO

	End If


End Function



Function GTFechador(GTWfecha)

	If CStr(GTWfecha & "") = "0" Or CStr(GTWfecha & "") = "" Then

		GTFechador  = "-----"

	Else



		GTWfec = CDate(GTWfecha + CLng(CDate("01-01-1900")))
		GTWndia = ofv.cdia(Weekday(GTWfec))
		GTWnMes = ofv.cmes(Month(GTWfec))

		GTwFechas = "  " & GTWndia & ", " & Day(GTWfec) & " de " & GTWnMes & " de " & Year(GTWfec)

		GTFechador  = UCase(GTwFechas)

	End If


End Function

Function GTFechadorsh(GTWfecha)

	If CStr(GTWfecha & "") = "0" Or CStr(GTWfecha & "") = "" Then

		GTFechadorSH  = "-----"

	Else


		GTwfec = CDate(GTWfecha + CLng(CDate("01-01-1900")))

		GTWFDIAB = CStr(Day(GTwfec))
		GTWFMESB = CStr(Month(GTwfec))

		If Len(GTWFDIAB) = 1 Then GTWFDIAB = "0" & GTWFDIAB
		If Len(GTWFMESB) = 1 Then GTWFMESB = "0" & GTWFMESB

		GTwFechash = GTWFDIAB & "/" & GTWFMESB & "/" & Year(GTwfec)
		GTFechadorSH  = UCase(GTwFechash)

	End If

End Function

Function GTCDSPHCALR(GTCDATEW,GTCDATEWL,GTCDATEWS)

	GTWhtm = ""

	GTWhtm = GTWhtm & " onclick="" " & Chr(13)
	' GTWhtm = GTWhtm & "   alert('" & GTCDATEW & "');" & chr(13)
	GTWhtm = GTWhtm & " DXCALDATE('" & GTCDATEW & "','" & GTCDATEWL & "','" & GTCDATEWS & "') ;" & Chr(13)
	GTWhtm = GTWhtm & """ " & Chr(13)

	GTCDSPHCALR = GTWhtm

End Function

Function GTCSETCALR(GTCFRM,GTCDAT,GTCSMT,GTCFCH,GTCFCHDAT,GTCIN,GTCYRAN,GTCYRDP,GTCMNAN,GTCMNDP,GTCTDIA)

	GTWhtm = ""

	GTWAccion = "ORCALENDAR.ASP?FCIN=" & GTCIN & "&FCYRAN=" & GTCYRAN & "&FCYRDP=" & GTCYRDP & "&FCMNAN=" & GTCMNAN & "&FCMNDP=" & GTCMNDP & "&FCTDIA=" & GTCTDIA & " "

	GTWhtm = GTWhtm & " onclick="" " & Chr(13)
	GTWhtm = GTWhtm & " var dxformw1 = document.getElementById('FRMCALENDAR'); " & Chr(13)
	GTWhtm = GTWhtm & " dxformw1.action = '" & GTWAccion & "'; " & Chr(13)
	GTWhtm = GTWhtm & " DXCDATO  = '" & GTCDAT & "';" & Chr(13)
	GTWhtm = GTWhtm & " DXCFORM  = '" & GTCFRM & "';" & Chr(13)

	If GTCSMT = "N" Then
		GTWhtm = GTWhtm & " DXCSMT  = 'N';" & Chr(13)
	End If

	If GTCFCH = "1"  Or GTCFCH = "2" Then
		GTWhtm = GTWhtm & " DXCFCH     = '" & GTCFCH & "';" & Chr(13)
		GTWhtm = GTWhtm & " DXCFCHDAT  = '" & GTCFCHDAT & "';" & Chr(13)
	End If



	GTWhtm = GTWhtm & " dxformw1.submit(); " & Chr(13)
	GTWhtm = GTWhtm & " document.getElementById('DXCALENDAR').style.display='';" & Chr(13)
	GTWhtm = GTWhtm & " return false;" & Chr(13)
	GTWhtm = GTWhtm & """ " & Chr(13)

	GTCSETCALR = GTWhtm

End Function

Function GTCCALENDAR(GTCTRG)

	If GTCTRG = "" Then
		GTCTRG = "fesquemas"
	ElseIf GTCTRG = "M" Then
		GTCTRG = "fesquemas"
	ElseIf GTCTRG = "D" Then
		GTCTRG = "WK4"
	End If

	GTWhtm = ""


	GTWhtm = GTWhtm & "<div id='DXCALENDAR' style='margin-top:1;margin-left:1;position:absolute;left:400px;top:200px;display:none;"
	GTWhtm = GTWhtm & "width:400;overflow:auto;"
	GTWhtm = GTWhtm & ""
	GTWhtm = GTWhtm & "' >"
	'''''' GTWhtm = GTWhtm & TABLA
	GTWhtm = GTWhtm & "</div>"

	GTWAccion = "XXXX"
	GTWhtm = GTWhtm & ofv.genform("FRMCALENDAR","aa style='margin:0;'",GTWAccion,GTCTRG,"post","","","si")



	GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & Chr(13)

	GTWhtm = GTWhtm & Chr(13)
	GTWhtm = GTWhtm & " var DXCDATO  = '';" & Chr(13)
	GTWhtm = GTWhtm & " var DXCFORM  = '';" & Chr(13)
	GTWhtm = GTWhtm & " var DXCSMT   = 'S';" & Chr(13)
	GTWhtm = GTWhtm & " var DXCFCH   = 'N';" & Chr(13)
	GTWhtm = GTWhtm & " var DXCFCHDAT  = '';" & Chr(13)
	GTWhtm = GTWhtm & Chr(13)
	GTWhtm = GTWhtm & "function DXCALDATE(DXCDATE,DXCDATEL,DXCDATES) {" & Chr(13)
	GTWhtm = GTWhtm & "    "
	GTWhtm = GTWhtm & "   var dxformw   = document.getElementById(DXCFORM);" & Chr(13)
	GTWhtm = GTWhtm & "   var dxdatow   =  document.getElementById(DXCDATO);" & Chr(13)
	''' ' GTWhtm = GTWhtm & "   var dxdatow   = document.c(DXCDATO);" & chr(13)
	GTWhtm = GTWhtm & " dxdatow.value=DXCDATE;" & Chr(13) 'alert(dxdatow.value);'

	GTWhtm = GTWhtm & "  if (DXCFCH=='N') {" & Chr(13)
	GTWhtm = GTWhtm & "     } else {" & Chr(13)
	GTWhtm = GTWhtm & "      var dxcellw   = document.getElementById(DXCFCHDAT);" & Chr(13)
	GTWhtm = GTWhtm & "      if (DXCFCH=='1') { " & Chr(13)
	GTWhtm = GTWhtm & "          dxcellw.innerHTML=DXCDATEL;" & Chr(13)
	GTWhtm = GTWhtm & "    } else {" & Chr(13)
	GTWhtm = GTWhtm & "          dxcellw.innerHTML=DXCDATES;" & Chr(13)
	GTWhtm = GTWhtm & "               }" & Chr(13)
	GTWhtm = GTWhtm & "                   }" & Chr(13)

	GTWhtm = GTWhtm & "  if (DXCSMT=='S') {dxformw.submit();}" & Chr(13)
	GTWhtm = GTWhtm & "  else {document.getElementById('DXCALENDAR').style.display='none';} " & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)

	GTWhtm = GTWhtm & "</SCRIPT>"

	GTCCALENDAR = GTWhtm


End Function

Function GTCOBS()

	GTWhtm = ""

	GTWhtm = GTWhtm & "<div id='DX4' style='position:absolute;left:10%;top:15%;display:none;"
	GTWhtm = GTWhtm & "width:600px;height:450px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
	GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
	GTWhtm = GTWhtm & "' >"

	GTWhtm = GTWhtm & "&nbsp;"

	GTWhtm = GTWhtm & "</div>"

	'GTWhtm = GTWhtm & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>"
	GTWhtm= GTWhtm &  "<script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & Chr(13)
	GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & Chr(13)
	GTWhtm = GTWhtm & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSWM() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='800px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSSS() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='500px';" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='350px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSHM() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='600px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)

	GTWhtm = GTWhtm & "function GTCOBSCLEX(e) {" & Chr(13)
	GTWhtm = GTWhtm & " if(e != 'S')  { alert('Por favor, complete los campos.');}  else { document.getElementById('DX4').style.display='none'; } " & Chr(13)

	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSCL() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.display='none';" & Chr(13)

	GTWhtm = GTWhtm & " } " & Chr(13)
	If session("yider")<> "" Then
		session("yider")=Replace(session("yider"),"*" ,"")
		ayider=Split(session("yider")," ")
		If IsArray(ayider) Then
			zdelim= ""
			pals=""
			For ia=0 To UBound(ayider)
				pal= Trim(ayider(ia) & "")
				If Len(pal) > 3 Then
					pals=pals & zdelim & pal
					zdelim= " "
				End If
			Next
			session("yider")=pals
		End If
		GTWhtm = GTWhtm & " localSearchHighlight('?h="& session("yider")& "'); var z= document.getElementById('sxxxs');" & Chr(13)
		GTWhtm = GTWhtm & "z.scrollIntoView(true);" & Chr(13)
	End If
	GTWhtm = GTWhtm & "</SCRIPT>" '& pals & "222222222222222222"

	GTCOBS = GTWhtm


End Function

Function GTCOBSRSK()

	GTWhtm = ""

	GTWhtm = GTWhtm & "<div id='DXA' style='position:absolute;left:50px;top:250px;display:none;"
	GTWhtm = GTWhtm & "width:500px;height:350px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
	GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
	GTWhtm = GTWhtm & "' >"

	GTWhtm = GTWhtm & "&nbsp;"

	GTWhtm = GTWhtm & "</div>"

	GTWhtm = GTWhtm & "<script type=" & Chr(34) & "text/javascript"& Chr(34) & " src='IncFile/dragdrop.js' ></script>"

	GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & Chr(13)
	GTWhtm = GTWhtm & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSWM() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DXA').style.width='800px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSSS() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DXA').style.width='500px';" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DXA').style.height='350px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSHM() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DXA').style.height='600px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCOBSCL() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DXA').style.display='none';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "</SCRIPT>"

	GTCOBSRSK = GTWhtm


End Function

Function GTCPTOFUN()

	GTWhtm = ""

	GTWhtm = GTWhtm & "<div id='DX4' style='position:absolute;left:200px;top:100px;display:none;"
	GTWhtm = GTWhtm & "width:700px;height:500px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
	GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
	GTWhtm = GTWhtm & "' >"

	GTWhtm = GTWhtm & "&nbsp;"

	GTWhtm = GTWhtm & "</div>"

	GTWhtm = GTWhtm & "<script type=" & Chr(34) & "text/javascript"& Chr(34) & " src='IncFile/dragdrop.js' ></script>"

	GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & Chr(13)
	GTWhtm = GTWhtm & Chr(13)
	GTWhtm = GTWhtm & "function GTCPTOFUNWM() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='800px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCPTOFUNSS() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='500px';" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='400px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "function GTCPTOFUNHM() {" & Chr(13)
	GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='600px';" & Chr(13)
	GTWhtm = GTWhtm & " } " & Chr(13)
	GTWhtm = GTWhtm & "</SCRIPT>"

	GTCPTOFUN = GTWhtm


End Function

Function ValidarInput()
	Dim sSval


	Sval= "<script languaje=javascript> "
	Sval=Sval & " function validarnum(obj,xform,vant,digitos) {"
	Sval= Sval & " var re=new RegExp('^\\d{'+ digitos +'}$'); "
	Sval= Sval & "if (obj.value.search(re)==-1) {alert('debe ser numerico y de ' + digitos +  ' posiciones'); obj.value=vant;return false;}"
	Sval = Sval & "else{  var fr= document.getElementById(xform); fr.submit();}}"

	Sval=Sval & " function validarnum2(obj,xform,vant,digitos) {"
	Sval= Sval & " var re=new RegExp('^\\d{'+ digitos +'}$'); "
	Sval= Sval & "if ((obj.value.search(re)==-1) && (obj.value.replace(' ','').length > 0) ) {alert('debe ser numerico y de ' + digitos +  ' posiciones' );return false;}"
	Sval = Sval & "else{  var fr= document.getElementById(xform);obj.value=obj.value.replace(' ','');fr.submit();}}"

	Sval= Sval & "</script>"
	validarinput=Sval

End Function


Function strPad( strToPad, toLength, padWith, padWhere )
	'====
	'padWhere: 1=right, 2=left, 3=both
	strToPad = strToPad & ""
	Dim padBoth
	padBoth = False

	If Not IsNumeric( toLength ) Then

	Else
		toLength =  toLength
	End If

	If padWith = "" Then padWith = " "

	If Not IsNumeric( padWhere ) Then
		padWhere = 1
	Else
		padWhere =  padWhere
		If padWhere > 3 Or padWhere < 1 Then padWhere = 1
		If padWhere = 3 Then
			padBoth = True
			padWhere = 1
		End If
	End If

	Do Until Len( strtoPad ) >= toLength
		Select Case padWhere
			Case 1
			strtoPad = strtoPad & padWith
			Case 2
			strtoPad = padWith & strtoPad
		End Select
		If padBoth Then
			If padWhere = 1 Then
				padWhere = 2
			Else
				padWhere = 1
			End If
		End If
	Loop

	If Len( strToPad ) > toLength Then
		If ( padBoth And padWhere = 1 ) Or ( padWhere = 2 And Not padBoth ) Then '=== no carriage return here
			strToPad = Right( strToPad, toLength )
		Else
			strToPad = Left( strToPad, toLength )
		End If
	End If

	strPad = strToPad

End Function


' devuelve la fecha inicio y fin de un trimestre

Function GetFecTrim(Trimestre,ano,fecini,fecfin)


	Select Case Trimestre
		Case 1
		ConvDate = DateSerial( ano,1,1)
		Case 2
		ConvDate = DateSerial( ano,4,1)
		Case 3
		ConvDate = DateSerial( ano,7,1)
		Case 4
		ConvDate = DateSerial( ano,10,1)
	End Select

	fecini=ConvDate
	fecfin = DateAdd("q", 1, fecini)
	fecfin = DateAdd("d", -1, fecfin)


End Function




' devuelve el trimestre de una fecha

Function GetTrimestre(fecha)

	dia=Day(fecha)
	mes=Month(fecha)
	ano=Year(fecha)
	fecha=DateSerial(ano,mes,dia)
	GetTrimestre=DatePart("q",fecha)

End Function

'------------------------------------------------------------

Function reemplaza(texto, patrn, replStr)
	Dim regEx
	Set regEx = New RegExp            ' Crea una expresi n regular.
	regEx.Pattern = patrn            ' Establece el modelo.
	regEx.IgnoreCase = True            ' Hace distinci n entre may sculas y min sculas.
	reemplaza = regEx.Replace(texto, replStr)   ' Realiza el reemplazo.
	Set regEx = Nothing
End Function


'------------------------------------------------------------

Function EliminarAcentos(TEXTO)

	If Len(texto) <> 0 Then
		texto = reemplaza(texto,"[      ]","A")
		texto = reemplaza(texto,"[      ]","E")
		texto = reemplaza(texto,"[      ]","I")
		texto = reemplaza(texto,"[      ]","O")
		texto = reemplaza(texto,"[      ]","U")
	End If
	EliminarAcentos = texto
End Function


Function GTCSETCALREX(GTCFRM, GTCDAT, GTCSMT, GTCFCH, GTCFCHDAT, GTCIN, GTCYRAN, GTCYRDP, GTCMNAN, GTCMNDP, GTCTDIA,GTCINIEX,GTCFINEX,GTCDIANTEX,GTCDIADPEX)

	GTWhtm = ""

	GTWAccion = "ORCALENDAR.ASP?FCIN=" & GTCIN & "&FCYRAN=" & GTCYRAN & "&FCYRDP=" & GTCYRDP & "&FCMNAN=" & GTCMNAN & "&FCMNDP=" & GTCMNDP & "&FCTDIA=" & GTCTDIA '& ""
	GTWAccion = GTWAccion & "&FDIASINI=" & GTCINIEX & "&FFIN=" & GTCFINEX & "&DIASANT=" & GTCDIANTEX & "&DIASDP=" & GTCDIADPEX
	GTWhtm = GTWhtm & " onclick="" " & Chr(13)
	GTWhtm = GTWhtm & " var dxformw1 = document.getElementById('FRMCALENDAR'); " & Chr(13)
	GTWhtm = GTWhtm & " dxformw1.action = '" & GTWAccion & "'; " & Chr(13)
	GTWhtm = GTWhtm & " DXCDATO  = '" & GTCDAT & "';" & Chr(13)
	GTWhtm = GTWhtm & " DXCFORM  = '" & GTCFRM & "';" & Chr(13)

	If GTCSMT = "N" Then
		GTWhtm = GTWhtm & " DXCSMT  = 'N';" & Chr(13)
	End If

	If GTCFCH = "1" Or GTCFCH = "2" Then
		GTWhtm = GTWhtm & " DXCFCH     = '" & GTCFCH & "';" & Chr(13)
		GTWhtm = GTWhtm & " DXCFCHDAT  = '" & GTCFCHDAT & "';" & Chr(13)
	End If

	GTWhtm = GTWhtm & " dxformw1.submit(); " & Chr(13)
	GTWhtm = GTWhtm & " document.getElementById('DXCALENDAR').style.display='';" & Chr(13)
	GTWhtm = GTWhtm & " return false;" & Chr(13)
	GTWhtm = GTWhtm & """ " & Chr(13)

	GTCSETCALREX = GTWhtm

End Function

'------------------------------------------------------------
Function BUSCARCAR(texto, patrn)
	Dim regEx
	Set regEx = New RegExp            ' Crea una expresi n regular.
	regEx.Pattern = patrn            ' Establece el modelo.
	regEx.IgnoreCase = True            ' Hace distinci n entre may sculas y min sculas.
	Set XBUSCAR = regEx.Execute(texto)   ' Realiza el reemplazo.
	BUSCARCAR=XBUSCAR.Count > 0
	Set XBUSCAR = Nothing
	Set regEx = Nothing
End Function


Function BUSCAEXCAR(TEXTO,PATRON)
	TEXTO=TEXTO &""
	If PATRON="" Then PATRON="["& Chr(9)  & Chr(34) & "'                              ]"

	BUSCAEXCAR=BUSCARCAR(TEXTO,PATRON)

End Function

Sub SortDictionary(ByRef dictionaryToSort, ByVal sortBy, ByVal orderType)
	If dictionaryToSort.Count>0 Then
		Dim k : k = dictionaryToSort.Keys
		Dim v : v = dictionaryToSort.Items
		Dim c : c = dictionaryToSort.Count - 1
		Dim e() : ReDim e(c, 1)
		'save dictionary in a multi-dimensional array
		For i = 0 To c
			e(i,0) = k(i) 'key name
			e(i,1) = v(i) 'key value
		Next
		'select sort type
		Dim t : t = 0
		Select Case sortBy
			Case SORTBY_KEY
			t = 0
			Case SORTBY_VALUE
			t = 1
		End Select
		'sort array by key or value then by ascending or descending order
		Dim i,j,tmp_k,tmp_v
		For i = 0 To c
			For j = i + 1 To c
				Select Case orderType
					Case ORDERTYPE_ASC
					If e(i,t) > e(j,t) Then
						tmp_k = e(i,0)
						tmp_v = e(i,1)
						e(i,0) = e(j,0)
						e(i,1) = e(j,1)
						e(j,0) = tmp_k
						e(j,1) = tmp_v
					End If
					Case ORDERTYPE_DESC
					If e(i,t) < e(j,t) Then
						tmp_k = e(i,0)
						tmp_v = e(i,1)
						e(i,0) = e(j,0)
						e(i,1) = e(j,1)
						e(j,0) = tmp_k
						e(j,1) = tmp_v
					End If
				End Select
			Next
		Next
		'empty dictionary and save sorted items
		dictionaryToSort.RemoveAll
		For i = 0 To c
			dictionaryToSort.Add e(i,0),e(i,1)
		Next
	End If

End Sub



Function armaperfiles(hash,cn,id)
	delim=""
	Dim sqlq
	HashPerfiles=" var  " & hash & "  =new Hashtable();" & Chr(13)

	sqlq="select a.descri,b.*,c.objid from Tipoperfil a "
	sqlq=sqlq & " inner JOIN ecoperfil b on( b.codperfil=a.tipoPERfil)"
	sqlq=sqlq & " inner join   objsector c on(c.codsector=b.codsector and c.codobjeto="  & id & ")"
	' response.write sqlq
	Set xpf=ofv.crearconsultaEx(sqlq,cn,1,parametros)
	' response.write sql
	Do While Not xpf.eof
		HashPerfiles= HashPerfiles & hash & ".put('pf" & xpf(1)	& "','" & xpf(0) & "'); " &  Chr(13)
		sperfiles= sperfiles & delim & "pf" & xpf(1) & "#" & xpf(0) & "#" & xpf(5)
		delim="###"
		xpf.movenext
	Loop
	Call ofv.cerrarconsulta(xpf)
	armaperfiles=HashPerfiles
End Function


Function armasectores(hash,cn,id)
	delim=""
	Dim sqlq
	HashPerfiles=" var  " & hash & "  =new Hashtable();" & Chr(13)

	delim=""
	sqlq=" Select b.descri,A.objid From objsector a "
	sqlq = sqlq & " inner join eco b on (b.eco=a.codsector)"

	sqlq =sqlq & " where a.codobjeto=0" & id
	'response.write sqlq
	Set xpf=ofv.crearconsultaEx(sqlq,cn,1,parametros)
	' response.write sql
	Do While Not xpf.eof
		HashPerfiles= HashPerfiles & hash & ".put('idsec" & xpf(1)	&"','" & xpf(0) &"'); " &  Chr(13)
		ssectores= ssectores & delim & "idsec" & xpf(1) & "#" & xpf(0)
		delim="###"

		xpf.movenext
	Loop
	Call ofv.cerrarconsulta(xpf)
	armasectores=HashPerfiles
End Function




Function retnum(s)

	For i =1 To Len(s)

		If IsNumeric(Mid(s,i,1)) Then
			retnum=retnum & (Mid(s,i,1))
		End If

	Next

End Function

Function EsqnomPx(xtipo, xcodigo, ofv)

	Set xcn = ofv.conectar(ofv.strconn2)


	If xtipo = "RSK" Then

		StrSql = "SELECT * FROM  ORRSKCIRCUITO "
		StrSql = StrSql & " Where OBJID = 0" & xcodigo
		Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
		If Not xrs.EOF Then
			If xrs("APLIC") = "OR" Then
				xdescri = "CIRCUITO AUTOEVALUACION"
			Else
				xdescri = "PLAN DE CONTINGENCIAS"
			End If

			StrSql = "Select * From objetos where codobjeto = 0" & xrs("CODOBJETO")

			Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
			If Not xrs.EOF Then
				xiden = xrs(1)
				xobjdes = xrs(3)
			End If


		End If
		Call ofv.cerrarconsulta(xrs)

		sTipobjetoX = (xdescri) & ": " & (xobjdes)

	Else


		sTipobjetoX = xcodigo
		sql="select * from tipobjeto where tipobjeto='" & xtipo &"'"
		'  RESPONSE.WRITE SQL
		Set RSM = ofv.crearconsultaEx(sql, xcn, 1, parametros)
		tipocpt=""
		If Not rsm.eof Then tipocpt=rsm("tipocpt")
		Call OFV.CERRARCONSULTA(RSM)


		'response.write tipocpt
		Select Case tipocpt


			Case "PD"
			StrSql2 = "Select a.* ,b.nombre From PEDIDOS a "
			StrSql2 = StrSql2 & " inner join pedidosdet b on (b.idpedido=a.objid) "
			StrSql2 = StrSql2  & " Where a.OBJID = 0" & xcodigo
			' response.write strsql2
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			If Not RSM.EOF Then
				XX=GETINFOUSER(RSM("SOLICITANTE"))
				sTipobjetoX= " Pedido Nro: " & RSM(0) & " De "  & SESSION("SGU_NOMBRE")
				sTipobjetoX =  sTipobjetoX &  "&nbsp;&nbsp;&nbsp;" & Left(RSM("nombre") ,100) & "..."
			End If
			Call OFV.CERRARCONSULTA(RSM)


			Case "SN"

			StrSql2 = "Select a.* ,b.nombre From PEDIDOS a "
			StrSql2 = StrSql2 & " inner join pedidosdet b on (b.idpedido=a.objid) "
			StrSql2 = StrSql2  & " Where a.OBJID = 0" & xcodigo
			' response.write StrSql2
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			If Not RSM.EOF Then
				XX=GETINFOUSER(RSM("SOLICITANTE"))
				sTipobjetoX= " Solicitud Nro: " & RSM(0) & " De "  & SESSION("SGU_NOMBRE")
				sTipobjetoX =  sTipobjetoX &  "&nbsp;&nbsp;&nbsp;" & Left(RSM("nombre") ,100) & "..."
			End If
			Call OFV.CERRARCONSULTA(RSM)

			Case Else


			If session("EMPMUN") = "S" And session("UN") = 1 Then

				StrSql2 = "Select a.*, c.cliente From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "
				StrSql2 = StrSql2 & "Where a.CODOBJETO = 0" & xcodigo

				StrSql2 = "Select a.*,c1.tipocpt,c.cliente,f.codigoform,B.OBJUN  From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "

				StrSql2 =StrSql2 &  "inner join tipobjeto c1 on ( c1.tipobjeto=a.tipobjeto) "
				StrSql2 = StrSql2  & " left join  objpropiedades f on (a.codobjeto=f.codobjeto)"
				StrSql2 = StrSql2  & "   Where a.CODOBJETO = 0" & xcodigo


			Else

				StrSql2 = "Select a.*, c.cliente From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "
				StrSql2 = StrSql2 & "Where a.CODOBJETO = 0" & xcodigo

				StrSql2 = "Select a.*,c1.tipocpt,c.cliente,f.codigoform,B.OBJUN From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "

				StrSql2 =StrSql2 &  "inner join tipobjeto c1 on ( c1.tipobjeto=a.tipobjeto) "
				StrSql2 = StrSql2  & " left join  objpropiedades f on (a.codobjeto=f.codobjeto)"
				StrSql2 = StrSql2  & "   Where a.CODOBJETO = 0" & xcodigo

			End If

			'response.write StrSql2 & "<br>"
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			IOBJUN=0
			If Not RSM.EOF Then
				IOBJUN=clng("0" & RSM("OBJUN"))
				tipo=RSM(2)


				If session("APLMDCLI") ="BCd" Then
					sTipobjeto1 = (RSM(1)) & ":&nbsp;"
					sTipobjeto2 = RSM(3) & ""
					sTipobjeto3 = " - " & (RSM(2) & "")
					sTipobjeto4 = " (" & (RSM(5) & "") & " "
					sTipobjeto5 = (RSM(6) & "") & ") "
				Else
				    atipod=split(rsm(1)," ")
				    if ubound(atipod) > 0 then
				   ' RESPOSNE.WRITE "222"
				      sTipobjeto1 = "<div class='item tipo-p'>"
				      sTipobjeto1 =  sTipobjeto1 & "<div class='badge' title='" &  rsm(1) & "' >" & ucase(left(atipod(0)& "",1))  &  ucase(left(atipod(1)& "",1)) & "</div>  "
   					  sTipobjeto1 =  sTipobjeto1 & "</div>"
				     ' sTipobjeto1 = (left(atipod(0)& "",1) ) & "" &  left(atipod(1)& "",1) & "  "
				    else
				     sTipobjeto1 = "<div class='item tipo-p'>"
					 sTipobjeto1 =  sTipobjeto1 & "<div class='badge' title='" & atipod(0) & "'>" & ucase(left(atipod(0)& "",1))  & "</div> </DIV> "

				    end if
					sTipobjeto1 = (RSM(1)) & ":&nbsp;"
					' sTipobjeto1=""
					If RSM("tipocpt")="K" Then
						'if RSM(8) = "9" then

						'else
						sTipobjeto1 =sTipobjeto1 & "&nbsp;"
						'end if
					End If
					sTipobjeto2 = (RSM(3) & "")

					sTipobjeto3 = " - " & (RSM(2) & "")
					sTipobjeto4 = " (" & (RSM(5) & "")
					If UCase(RSM(5) & "")="O" Then
						sTipobjeto5 = ")"
					Else
						sTipobjeto5 =  " " & (RSM(6) & "") & ") "
					End If
				End If
				If session("FULLDOCN") = "1" And session("APLMDCLI")="BGAL" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto4 & sTipobjeto5
				ElseIf  session("FULLDOCN") = "1" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2
				ElseIf session("FULLDOCN") = "2" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
				ElseIf session("FULLDOCN") = "4" Then
					sTipobjetoX =  sTipobjeto1 & tipo & " - " & sTipobjeto2 &  sTipobjeto4 & sTipobjeto5
				ElseIf session("FULLDOCN") = "5" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto4 & sTipobjeto5
				Else
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3  & sTipobjeto4 & sTipobjeto5
				End If



			End If


		End Select

		'NUEVO 02/08/2021 GP --> ALE
		'10317-0004/20
		Select Case session("APLMDCLI")

			Case "GP"
            if  session("UN")="" then  session("UN")=1
			SQLAC="SELECT CODEMPRE FROM OBJMEMPALC WHERE TIPOBJETO='"  & xtipo &  "' AND CODOBJETO=" & xcodigo & " AND CODEMPRE=" & session("UN")
			 ' response.write sqlac
			Set RSMALC = ofv.crearconsultaEx(SQLAC, xcn, 1, parametros)

			STRCLIENTE="NO"
			If Not RSMALC.EOF Then
				STRCLIENTE="SI"
			End If
			Call OFV.CERRARCONSULTA (RSMALC)

			If session("EMPMUN") = "S" And (session("UN") = 1  Or (IOBJUN = 1 And STRCLIENTE="SI" )) And tipocpt <> "PD" And tipocpt <> "SN" Then

				EsqnomPx = sTipobjetoX & "&nbsp;-&nbsp;" & cstr(RSM("cliente") &"")
			Else
				EsqnomPx = sTipobjetoX
			End If

			Case Else

			If session("EMPMUN") = "S" And session("UN") = 1 And tipocpt <> "PD" And tipocpt <> "SN" Then
				If Not RSM.eof Then
					EsqnomPx = sTipobjetoX & "&nbsp;-&nbsp;" & cstr(RSM("cliente") &"")
				Else
					EsqnomPx = sTipobjetoX
				End If
			Else
				EsqnomPx = sTipobjetoX
			End If

		End Select

	End If





   ' EsqnomPx=sTipobjeto1 & "<div class='item__texto'>" & EsqnomPx & "</div>"


	Call OFV.CERRARCONSULTA (RSM)
	Call ofv.cerrarconn(xcn)

End Function



Function EsqnomPxEx(xtipo, xcodigo, ofv,dato)

	Set xcn = ofv.conectar(ofv.strconn2)


	If xtipo = "RSK" Then

		StrSql = "SELECT * FROM  ORRSKCIRCUITO "
		StrSql = StrSql & " Where OBJID = 0" & xcodigo
		Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
		If Not xrs.EOF Then
			If xrs("APLIC") = "OR" Then
				xdescri = "CIRCUITO AUTOEVALUACION"
			Else
				xdescri = "PLAN DE CONTINGENCIAS"
			End If

			StrSql = "Select * From objetos where codobjeto = 0" & xrs("CODOBJETO")

			Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
			If Not xrs.EOF Then
				xiden = xrs(1)
				xobjdes = xrs(3)
			End If


		End If
		Call ofv.cerrarconsulta(xrs)

		sTipobjetoX = (xdescri) & ": " & (xobjdes)

	Else


		sTipobjetoX = xcodigo
		sql="select * from tipobjeto where tipobjeto='" & xtipo &"'"
		'  RESPONSE.WRITE SQL
		Set RSM = ofv.crearconsultaEx(sql, xcn, 1, parametros)
		tipocpt=""
		If Not rsm.eof Then tipocpt=rsm("tipocpt")
		Call OFV.CERRARCONSULTA(RSM)


		'response.write tipocpt
		Select Case tipocpt


			Case "PD"
			StrSql2 = "Select a.* ,b.nombre From PEDIDOS a "
			StrSql2 = StrSql2 & " inner join pedidosdet b on (b.idpedido=a.objid) "
			StrSql2 = StrSql2  & " Where a.OBJID = 0" & xcodigo
			' response.write strsql2
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			If Not RSM.EOF Then
				XX=GETINFOUSER(RSM("SOLICITANTE"))
				sTipobjetoX= " Pedido Nro: " & RSM(0) & " De "  & SESSION("SGU_NOMBRE")
				sTipobjetoX =  sTipobjetoX &  "&nbsp;&nbsp;&nbsp;" & Left(RSM("nombre") ,100) & "..."
			End If
			Call OFV.CERRARCONSULTA(RSM)


			Case "SN"

			StrSql2 = "Select a.* ,b.nombre From PEDIDOS a "
			StrSql2 = StrSql2 & " inner join pedidosdet b on (b.idpedido=a.objid) "
			StrSql2 = StrSql2  & " Where a.OBJID = 0" & xcodigo
			' response.write StrSql2
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			If Not RSM.EOF Then
				XX=GETINFOUSER(RSM("SOLICITANTE"))
				sTipobjetoX= " Solicitud Nro: " & RSM(0) & " De "  & SESSION("SGU_NOMBRE")
				sTipobjetoX =  sTipobjetoX &  "&nbsp;&nbsp;&nbsp;" & Left(RSM("nombre") ,100) & "..."
			End If
			Call OFV.CERRARCONSULTA(RSM)

			Case Else


			If session("EMPMUN") = "S" And session("UN") = 1 Then

				StrSql2 = "Select a.*, c.cliente From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "
				StrSql2 = StrSql2 & "Where a.CODOBJETO = 0" & xcodigo

				StrSql2 = "Select a.*,c1.tipocpt,c.cliente,f.codigoform,B.OBJUN  From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "

				StrSql2 =StrSql2 &  "inner join tipobjeto c1 on ( c1.tipobjeto=a.tipobjeto) "
				StrSql2 = StrSql2  & " left join  objpropiedades f on (a.codobjeto=f.codobjeto)"
				StrSql2 = StrSql2  & "   Where a.CODOBJETO = 0" & xcodigo


			Else

				StrSql2 = "Select a.*, c.cliente From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "
				StrSql2 = StrSql2 & "Where a.CODOBJETO = 0" & xcodigo

				StrSql2 = "Select a.*,c1.tipocpt,c.cliente,f.codigoform,B.OBJUN From CSOBJETOSRDO a "
				StrSql2 = StrSql2 & " inner join objetos b on (a.codobjeto=b.codobjeto) "
				StrSql2 = StrSql2 & " inner join clienteplan c on (b.objun=c.id) "

				StrSql2 =StrSql2 &  "inner join tipobjeto c1 on ( c1.tipobjeto=a.tipobjeto) "
				StrSql2 = StrSql2  & " left join  objpropiedades f on (a.codobjeto=f.codobjeto)"
				StrSql2 = StrSql2  & "   Where a.CODOBJETO = 0" & xcodigo

			End If

			'response.write StrSql2 & "<br>"
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			IOBJUN=0
			If Not RSM.EOF Then
				IOBJUN=clng("0" & RSM("OBJUN"))
				tipo=RSM(2)


				If session("APLMDCLI") ="BCw" Then
					sTipobjeto1 = (RSM(1)) & ":&nbsp;"
					sTipobjeto2 = RSM(3) & ""
					sTipobjeto3 = " - " & (RSM(2) & "")
					sTipobjeto4 = " (" & (RSM(5) & "") & " "
					sTipobjeto5 = (RSM(6) & "") & ") "
				Else
				    atipod=split(rsm(1)," ")
				    if ubound(atipod) > 0 then
				   ' RESPOSNE.WRITE "222"
				      sTipobjeto1 = "<div class='item tipo-p'>"
				      sTipobjeto1 =  sTipobjeto1 & "<div class='badge' title='" &  rsm(1) & "' >" & ucase(left(atipod(0)& "",1))  &  ucase(left(atipod(1)& "",1)) & "</div>  "
   					'  sTipobjeto1 =  sTipobjeto1 & "</div>"
				     ' sTipobjeto1 = (left(atipod(0)& "",1) ) & "" &  left(atipod(1)& "",1) & "  "
				     '	sTipobjeto1 = (RSM(1)) & ":&nbsp;"
				    else

				     sTipobjeto1 = "<div class='item tipo-p'>"
					 sTipobjeto1 =  sTipobjeto1 & "<div class='badge' title='" & atipod(0) & "'>" & ucase(left(atipod(0)& "",1))  & "</div>"
                     '  sTipobjeto1 = (RSM(1)) & ":&nbsp;"
				    end if
					sTipobjeto1 = (RSM(1)) & ":&nbsp;"
					' sTipobjeto1=""
					If RSM("tipocpt")="K" Then
						'if RSM(8) = "9" then

						'else
						sTipobjeto1 =sTipobjeto1 & "&nbsp;"
						'end if
					End If
					sTipobjeto2 = (RSM(3) & "")

					sTipobjeto3 = " - " & (RSM(2) & "")
					sTipobjeto4 = " (" & (RSM(5) & "")
					If UCase(RSM(5) & "")="O" Then
						sTipobjeto5 = ")"
					Else
						sTipobjeto5 =  " " & (RSM(6) & "") & ") "
					End If
				End If


				select case dato
				case "1"
				  sTipobjetoX = sTipobjeto2 & sTipobjeto3  & sTipobjeto4 & sTipobjeto5
				case else
				If session("FULLDOCN") = "1" And session("APLMDCLI")="BGAL" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto4 & sTipobjeto5
				ElseIf  session("FULLDOCN") = "1" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2
				ElseIf session("FULLDOCN") = "2" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
				ElseIf session("FULLDOCN") = "4" Then
					sTipobjetoX =  sTipobjeto1 & tipo & " - " & sTipobjeto2 &  sTipobjeto4 & sTipobjeto5
				ElseIf session("FULLDOCN") = "5" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto4 & sTipobjeto5
				Else
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3  & sTipobjeto4 & sTipobjeto5
				End If
               end select


			End If


		End Select

		'NUEVO 02/08/2021 GP --> ALE
		'10317-0004/20
		Select Case session("APLMDCLI")

			Case "GP"
            if  session("UN")="" then  session("UN")=1
			SQLAC="SELECT CODEMPRE FROM OBJMEMPALC WHERE TIPOBJETO='"  & xtipo &  "' AND CODOBJETO=" & xcodigo & " AND CODEMPRE=" & session("UN")
			 ' response.write sqlac
			Set RSMALC = ofv.crearconsultaEx(SQLAC, xcn, 1, parametros)

			STRCLIENTE="NO"
			If Not RSMALC.EOF Then
				STRCLIENTE="SI"
			End If
			Call OFV.CERRARCONSULTA (RSMALC)

			If session("EMPMUN") = "S" And (session("UN") = 1  Or (IOBJUN = 1 And STRCLIENTE="SI" )) And tipocpt <> "PD" And tipocpt <> "SN" Then

				EsqnomPxEx = sTipobjetoX & "&nbsp;-&nbsp;" & cstr(RSM("cliente") &"")
			Else
				EsqnomPxEx = sTipobjetoX
			End If

			Case Else

			If session("EMPMUN") = "S" And session("UN") = 1 And tipocpt <> "PD" And tipocpt <> "SN" Then
				If Not RSM.eof Then
					EsqnomPxEx = sTipobjetoX & "&nbsp;-&nbsp;" & cstr(RSM("cliente") &"")
				Else
					EsqnomPxEx = sTipobjetoX
				End If
			Else
				EsqnomPxEx = sTipobjetoX
			End If

		End Select

	End If





   ' EsqnomPxEx=sTipobjeto1q & "<div class='item__texto'>" & EsqnomPxEx & "</div>"
    EsqnomPxEx=EsqnomPxEx


	Call OFV.CERRARCONSULTA (RSM)
	Call ofv.cerrarconn(xcn)

End Function




Function empresas()

	EMPRESASX=""
	DELIM = ""
	'response.write SESSION("TAEMP")
	WATREMPTB = Split(SESSION("ATREMP"),Chr(10))

	For IAEMP=0 To SESSION("TAEMP")-1
		'response.write WATREMPTB(IAEMP) & "--<br>"
		WATREMPEL = Split(WATREMPTB(IAEMP),Chr(9))

		If WATREMPEL(2) = 1 Then
			'response.write  WATREMPEL(0) & "<br>"
			EMPRESASX = EMPRESASX & DELIM & WATREMPEL(0)
			DELIM = Chr(9)

		End If

	Next

	empresas = empresasX
End Function



Function ARMAM2()
	Dim cbmm

	Dim str
	Dim AUsrG
	str="" & Chr(13)
		XMESAURL = ""

		If session("URLTYP") = "V" Then
			XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
		End If


		If session("URLPORT") = "N" Then
			XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
		Else
			XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
		End If



	surx=XMESAURL & "/coremanager/novedmenu.asp"


    sinicio ="<div align=left xclass='item-container' xonclick=cerrar()><i class='la la-bars'></i></div><a class=ntmnn href='" &  surx & "' target=info><i class='las la-home'></i><font xface=calibri> INICIO</font></a>"
    sinicio ="<a class=ntmnn href='" &  surx & "' target=info><i class='las la-home'></i><font xface=calibri>INICIO</font></a>"


		xx=getinfouser(Session("Usuario"))
		suserinf= "Email:" & capitalize(UCase(Session("sgu_mail")& "" ))
		suserinf=suserinf & "  Grupos:"

		wktgrupos = Split(session("SegGrpdsc"), Chr(9))

		For wkti = 0 To UBound(wktgrupos)
			suserinf=suserinf & " " & capitalize(wktgrupos(wkti))
		Next

       suserinf=suserinf & " -> PERFIL=" & Session("UsrPstDsc")
	  'ARMAM2="<div id=contm xclass=sidebar><i class='la la-bars'></i><ul STYLE= 'overflow:auto;Height:100%;' id=myMenu >" & chr(13)
	   ARMAM2="<div id=contm STYLE='overflow:hidden;' xclass=sidebar><ul class=ulmenu STYLE= 'overflow:auto;Height:100%;' id=myMenu >" & chr(13)
	   'ARMAM2=ARMAM2 & chr(13) & "<li><div class='item-container'><div class='item-text'><a  class=ntmnn   href=# class='link'  >  <font xsize=2 xcolor=  #fdfefe   xface=calibri>" & sinicio & "</font></a>" & Chr(13) & "</div></div></li>" & Chr(13)
	   ARMAM2=ARMAM2 & chr(13) & "<li><div class='item-container'><div class='item-text'><img src='" & Session("CliLogo") &"'></div></div></li>" & Chr(13)


	 ' suserinf=suserinf  & "<i style=' font-size: 35px; ' class='la la-user'></i>  <font style=' font-size: 20px;'    face=calibri><i>" &  capitalize(Session("SegUsuario"))  & "</i></font><br>"


	  ARMAM2=ARMAM2 & chr(13) & "<li ><div class='item-container userm'><div class='item-text'><a class='ntmnn userc'  href=# class='link'  >  <i style=' font-size: 25px; '  class='la la-user'></i> <font  title='"& suserinf &"' xsize=2 xcolor=  #fdfefe   xface=calibri>" &  Session("SegUsuario") & "<br>" & xsuserinf& "</font></a>" & Chr(13) & "</div></div></li>" & Chr(13)


	If  Session("DELLOG") = "S" Then
		'  session("SegGrpLst")="MAS" & CHR(9) & "XX"
		AUsrG = Split(session("SegGrpLst"),Chr(9))
		AUsrG2 = Split(session("SegGrpLst"),Chr(9))
		XMFSEP=""
		For XMFI = 0 To UBound(AUsrG)
			UsrGS  = UsrGS  & XMFSEP & "'" & Replace(AUsrG(XMFI) & "","'","''") & "'"
			XMFSEP  = ","
		Next
	Else
		UsrGS=  "'" & Session("seggrp")  & "'"
		AUsrG2 = Split(Session("seggrp") ,Chr(9))

	End If

	Set cmm = OFV.conectar(OFV.strconn0)
	sub1=""

	SQL= "Select max (BB.codapli) as codapli ,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr "
	SQL=SQL & " from ( "
	XMFSEP=""
	SQL2=""
	For XMFI = 0 To UBound(AUsrG2)
		UsrGS  ="'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"
		sql2=sql2 & chr(13) & XMFSEP & chr(13)
		sql2=sql2 & "Select b.codapli as codapli ,isnull(  CAST(c.grpestado AS INT)  ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1) as atrUsr"
		sql2=sql2 &  " FROM CorAplic b "
		sql2=sql2 &  " left join atrmXgrupo c on   ( c.Tipo = 'A' and c.grupo IN("& UsrGS &") and  b.codapli=C.APLICACION) "
		sql2=sql2 &  " left join atrmXusuario d on   ( D.Tipo = 'A' and d.usuario='" & Session("segperfil") & "' and D.APLICACION=b.codapli) "
		sql2=sql2 &  " where b.enable <> 0"
		XMFSEP= CHR(13) & " UNION " & CHR(13)
	Next
	SQL =SQL & SQL2 & ") BB"
	SQL=SQL & " GROUP BY BB.codapli " & CHR(13) & CHR(13)

   session("SQL")=sql
	Set rsApp = ofv.crearconsultaEx(sql,cmm,1,parametros)
	AppH=""
	XMFSEP=""
	Do While Not rsapp.eof
		If rsAPP("atrGrupo") <> 0 And rsAPP("atrUsr") <> 0 Then
			AppH  = AppH  & XMFSEP & "'" & rsapp(0) & "'"
			XMFSEP  = ","
		End If

		rsapp.movenext
	Loop
	Call ofv.cerrarconsulta(rsapp)

	ARMAM2= ARMAM2 &  listarM2ex("0",0,cmm,AppH,UsrGS)

    suserinf=  "<i style=' font-size: 35px; ' class='la la-user'></i>  <font style=' font-size: 20px;'    face=calibri><i>" &  capitalize(Session("SegUsuario"))  & "</i></font><br>"

   ' ARMAM2=ARMAM2 & chr(13) & "<li ><div class='item-container'><div class='item-text spcapp'>MI PERFIL</div></div></li>" & Chr(13)
   ' ARMAM2=ARMAM2 & chr(13) & "<li ><div class='item-container'><div class='item-text'>" & suserinf & "</div></div></li>" & Chr(13)

	ARMAM2=ARMAM2 & "</ul></div>"

End Function

Function  ListarM2ex(mdepen,nivel,cmm,aplic,UsrgS)
	Dim rs,rs1
	Dim rsoef





	sql="Select max(bb.id), bb.MENU,bb.mcodapli,bb.titapli,max(bb.menuatr) as menuatr,MAX(isnull(  CAST(bb.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(bb.atrUsr AS INT) ,1) )as atrUsr, max(bb.morden) as morden,max(bb.ordenapli )  as ordenapli,max(bb.asp )  as asp"
	sql =sql & " from ("

	XMFSEP=""
	SQL2=""
	For XMFI = 0 To UBound(AUsrG2)
		UsrGS  ="'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"
		sql2=sql2 & chr(13) & XMFSEP & chr(13)
		sql2=sql2 &  "Select a.id, a.MENU,a.mcodapli,b.titapli,a.menuatr as menuatr,isnull(  CAST(c.grpestado AS INT)  ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1) as atrUsr, a.morden as morden,ordenapli   as ordenapli,asp  as asp "

		sql2=sql2 & "From Menux a "
		sql2=sql2 & " inner join CorAplic b on ( b.codapli = a.mcodapli ) "
		sql2=sql2 & " left join atrmXgrupo c on   (c.alias=a.menuatr and c.Tipo = 'S' and c.grupo IN("& UsrGS & ") and c.aplicacion=b.codapli) "
		sql2=sql2 & " left join atrmXusuario d on   (d.alias=a.menuatr and D.Tipo = 'S' and d.usuario='" & Session("segperfil") & "' and D.aplicacion=b.codapli) "
		sql2=sql2 & " where  b.CodApli in(" & aplic & " ) AND a.mdepen = '0'"
		' no trae formularios
		If session("CORMODUSOLABM") <> "S" Then
			sql2=sql2 & " and menuatr not in('CLIS245159671557361072846976','CLIS45312740033826860617908')"
		End If
		XMFSEP= CHR(13) & " UNION " & CHR(13)
	Next
	SQL =SQL & SQL2 & ") BB"
	sql =sql & "  GROUP BY bb.titapli,bb.mcodapli,bb.MENU"
	'sql =sql & "  order by  ordenapli,bb.titapli,bb.mcodapli,bb.MENU,morden"
	sql =sql & "  order by ordenapli,morden"

	session("SQL")=sql
	'   response.write sql & "<br>"

	Set rs = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If  rs.eof Then
		rseof="S"
	Else
		rseof="N"
	End If

	Dim a(100,1) , i
	Dim imenu , iSm
	' Recorre los subdirectorios
	i=0
    appant=""
    If Session("UsrCorp") And Session("EMPMUN") = "S" Then
    	XXXX = NEWEMPRE("NO")

		sEMPRE = request.querystring("EMPRE")

		IF SEMPRE = ""  AND SESSION("UN") = "" THEN

			SEMPRE = ATREMP(0, 0)
			SESSION("UN") = SEMPRE & ""
			If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
		elseIF sEMPRE <> "" THEN
			XXXX = NEWEMPRE(sEMPRE)
			SESSION("UN")=SEMPRE
		ELSE
			XXXX = NEWEMPRE(SESSION("UN"))
			sempre=SESSION("UN")
	end if
	contenido=""
    FOR IAEMP=1 TO TAEMP
			IF ATREMP(IAEMP,2) = 1 THEN
            sHref = "../coress20/novedmenuEstilo.asp?coraplic=COREM20&EMPRE=" & ATREMP(IAEMP,0)
            if ATREMP(IAEMP,0)=session("un") then
             contenido =contenido  & "<font><b><i><a style='padding-left:14px;' class='selec' title='" & ATREMP(IAEMP,4)& "' href='" & sHref & "'>" & ATREMP(IAEMP,4) & " </a> </i></b></font><br>"
            else
			 contenido =contenido  & "<font><a style='padding-left:14px;' class='selec' title='" & ATREMP(IAEMP,4)& "' href='" & sHref & "'>" & ATREMP(IAEMP,4) & " </a> </font><br>"
            end if
			end if
	next


    	str=str & chr(13) & "<li data-id='subm' class='lixstk' TITLE='' ><div class='item-container'><div class='item-text2' style='' ><a  class=ntmnn3 style=''  title='' href=# class='link'  data-id='' > <i style=' font-size: 25px; ' xclass='la la-pencil'></i> <font xsize=2 xcolor=  #fdfefe   xface=calibri>EMPRESAS<br>" & contenido & "</font></a>" & Chr(13) & "</div></div>" & Chr(13)
    end if
	Do While Not rs.eof
		'  ReDim Preserve a(i,0)

		If rs("atrGrupo") <> 0 And rs("atrUsr") <> 0 Then


		 if appant <>  rs(2) then
		   if appant="" then
		    str=str & chr(13) & "<li ><div class='item-container'><div class='item-text spcapp2'>" & Capitalize((rs("titapli"))) & " </div></div></li>" & Chr(13)
		   else
		    str=str & chr(13) & "<li ><div class='item-container'><div class='item-text spcapp'>" & Capitalize((rs("titapli"))) & "</div></div></li>" & Chr(13)
		   end if
		   appant=rs(2)
		 end if

			'if true then
			sql2="Select top 1  Menux.* From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & rs("menuatr") & "'"
			Set rs2 = ofv.crearconsultaEx(sql2,cmm,1,parametros)

			If Not rs2.eof Then
				'   response.write sql2 & chr(13)
				If rs(2)="COREM20" Then
					str=str & chr(13) & "<li data-id='subm" & rs(0) & "' class='lixstk' TITLE='" & rs("titapli") &"' ><div class='item-container'><div class='item-text'><a  class=ntmnn  title='" & rs("titapli") & "' href=# class='link'  data-id='subm" & rs(0) & "' > <i style=' font-size: 25px; ' xclass='la la-pencil'></i> <font xsize=2 xcolor=  #fdfefe   xface=calibri>" & Capitalize(UCase(rs("menu"))) & "</font></a>" & Chr(13) & "</div></div>" & Chr(13)
				Else
					str=str & chr(13) & "<li data-id='subm" & rs(0) & "'  class='lixstk' TITLE='" & rs("titapli") &"' ><div class='item-container'><div class='item-text'><a  class=ntmnn  title='" & rs("titapli") & "' href=# class='link'  data-id='subm" & rs(0) & "' > <i style=' font-size: 25px; ' xclass='la la-gear'></i> <font xsize=2 xcolor=  #fdfefe   xface=calibri>" & Capitalize(UCase(rs("menu"))) & "</font></a>" & Chr(13) & "</div></div>" & Chr(13)
				End If

				'sub menu
				'if session("SGU_userid")="CL" then
                 str=str & armasubmNEW2(rs("menuatr"),cmm,rs(1),1,"'" & rs(2) & "'" ,rs(0),UsrGS) & "</li>"
				'sub1=sub1 & "<div data-color='"& session("APLITBGD") & "' class='nnn' id=subm"  & rs(0)  & " style='display:none;'> " & armasubmNEW(rs("menuatr"),cmm,rs(1),1,"'" & rs(2) & "'" ,rs(0),UsrGS) & "</DIV>"
				'else
				'sub1=sub1 & armasubm(rs("menuatr"),cmm,rs(1),1,"'" & rs(2) & "'" ,rs(0),UsrGS)
				'end if

			Else
				xsel="menuatr"
				sInput = Trim(rs(xsel))
				sHref = Trim(rs("asp"))
				sTarget = "info"
				sPath = Session("CorPath")
				If rs(3) <> "" Then sTarget = rs(3)
				spt = InStr(rs("asp"), ".asp")
				xesm = "N"
				If spt > 1 Then
					If Left(rs("asp"), spt - 1) = "Submenus" Then
						sPath = "../coress20/"
						xesm = "S"
					End If
				End If
				sseg = InStr(sHref, "?")
				If sseg = 0 Then
					sHref = sHref & "?"
				Else
					sHref = sHref & "&"
				End If
				sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")
				ssMnn = "MNN" & OFV.convertircar(UCase(rs(1)), " ", "_")
				str=str & chr(13)  & "<LI class=lixstk><div class='item-container'><div class='item-text'><font xcolor=  #fdfefe   xsize=1 xface=calibri ><a onclick=" & Chr(34) & "SetMnn('" & ssMnn & "');" & Chr(34)  & " class=xxxx target=info href='" & sHref & "'>" & UCase(RS("MENU")) & "</a></font></div></div></LI>"  & Chr(13)
			End If
		End If


		rs.movenext

	Loop
	Call ofv.cerrarconsulta(rs)

	ListarM2ex=str
End Function


Function armasubmNEW2(id,cmm,idx,nivel,aplic,idx2,UsrGS)
	Dim ss7
	'dim AUsrG
	'dim rsx
	'dim XMFI

	'sql="Select Menux.*,dirapli,titapli From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & id & "' "
	SQL = "Select max(BB.id) as id, BB.MENU,BB.mcodapli,BB.titapli,max(BB.menuatr) as menuatr,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr, max(BB.morden) as morden,max(BB.ordenapli )  as ordenapli,max(BB.asp )  as asp"
	SQL= SQL & ",max(BB.target) as target,max(BB.dirapli) as dirapli"
	sql=sql  & " FROM ( " & chr(13)

	'   if  Session("DELLOG") = "S" then
	'    AUsrG = Split(session("SegGrpLst"),Chr(9))
	'   else
	'   Session("seggrp")=Session("seggrp") & chr(9) & Session("seggrp") & "222"
	'    AUsrG = Split(Session("seggrp"),Chr(9))
	'   end if

	XMFSEP=""
	sql2=""
	For XMFI = 0 To UBound(AUsrG2)
		UsrGS  =  "'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"

		sql2=sql2 & chr(13) & XMFSEP & chr(13)
		sql2=sql2 & "	Select a.id as id, a.MENU,a.mcodapli,b.titapli,a.menuatr as menuatr,isnull(  c.grpestado   ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1)as atrUsr, a.morden as morden,ordenapli   as ordenapli,asp  as asp"
		sql2=sql2 & "	, Dirapli  ,target as target"
		sql2=sql2 & "	 From Menux a "
		sql2=sql2 & "	inner join CorAplic b on ( b.codapli = a.mcodapli ) "
		sql2=sql2 & "	left join atrmXgrupo c on   ( c.tipo='M' and  c.alias=a.menuatr  and c.grupo IN("& UsrGS &") and c.aplicacion=b.codapli) "
		sql2=sql2 & "	left join atrmXusuario d on ( d.tipo='M' AND d.alias=a.menuatr  and d.usuario='"& Session("segperfil") &"' and D.aplicacion=b.codapli)"
		sql2=sql2 & "     where  b.CodApli in( " & aplic &" ) AND a.mdepen = '" & id& "'"
		'sql2= chr(13) &  XMFSEP & chr(13) & sql2
		XMFSEP  = "UNION"
	Next
	sql= sql & sql2 & chr(13) & " )BB " & chr(13)
	sql=sql & " GROUP BY BB.titapli,BB.mcodapli,BB.MENU"
	sql=sql & " order by  morden" & chr(13)  & chr(13)


	'ss7=""
	'if ID="SEGURIDAD" THEN
	'session("SQL")=SQL
	'END IF

	Set rsx = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If Not rsx.eof Then
			 ss7=ss7 & chr(13) & "<ul>"

		Do While Not rsx.eof

			' response.write "222"
			If rsx("atrGrupo") <> 0 And rsx("atrUsr") <> 0 Then
				STEXTO1=UCase(rsx("menu"))
				xsel="menuatr"
				sInput = Trim(rsx(xsel))
				XMESAURL = ""

				If session("URLTYP") = "V" Then
					XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
				End If


				If session("URLPORT") = "N" Then
					XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
				Else
					XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
				End If
				surx=	XMESAURL
				sHref = surx & "/" & rsx("dirapli") & "/" & Trim(rsx("asp"))
				sTarget = "info"
				sPath = rsx("dirapli")
				If rsx("target") <> "" Then sTarget = rsx("target")
				spt = InStr(rsx("asp"), ".asp")
				xesm = "N"
				If spt > 1 Then
					If Left(rsx("asp"), spt - 1) = "Submenus" Then
						sPath = "../coress20/"
						xesm = "S"
					End If
				End If
				sseg = InStr(sHref, "?")
				If sseg = 0 Then
					sHref = sHref & "?"
				Else
					sHref = sHref & "&"
				End If

				sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rsx("menu")), " ", "_")
				If instr(lcase(rsx("asp")),"base=") = 0 Then
					' agrego base=D para los menus que no son de formularios
					sHref = sHref &  "&base=D"
				End If
                sHref = sHref &  "&coraplic=" & rsx("mcodapli")
				ssMnn = "MNN" & OFV.convertircar(UCase(rsx("menu")), " ", "_")

				sAccion= sHref
				If nivel=1 Then
					If instr(lcase(sAccion),"submenus.asp") > 0 Then
						ss7 = ss7  & "<LI class='spc' TITLE='" & RSX("titapli")& "'><div class='item-container'><div class='item-text'><font xclass=spc3  xcolor=  #fdfefe   xsize=1 xface=verdana ><a class='ntmnn spcex' xtarget=info xhref='"& sAccion & "'>"  & Capitalize(Replace(STEXTO1," ","&nbsp;")) & "</a></font></DIV></DIV>" '"<li>" &  stexto1  & chr(13)
					    ss7=ss7 & armasubmNEW2(rsx("menuatr"),cmm,rsx("menu"),nivel +1,aplic,rsx("id"),UsrGS)

					Else
						ss7 = ss7  & "<LI class=spc TITLE='" & RSX("titapli")& "'><div class='item-container'><div class='item-text'><font xclass=spc3  xcolor=  #fdfefe   xsize=1 xface=verdana ><a class='ntmnn spcex' xtitle='" & rsx("titapli") & "'target=info href='"& sAccion & "'>"  & Capitalize(Replace(STEXTO1," ","&nbsp;")) & "</a></font></DIV></DIV>" '"<li>" &  stexto1  & chr(13)
					End If

				Else
					ss7 = ss7  & "<LI  class=spc  TITLE='" & RSX("titapli")& "' ><div class='item-container'><div class='item-text'><font xclass=spc3  xcolor=  #fdfefe   xsize=1 xface=verdana ><a class='ntmnn spcex2' xtitle='" & rsx("titapli") & "' target=info href='"& sAccion & "'>&nbsp;&nbsp;&nbsp;" &  Capitalize(Replace(STEXTO1," ","&nbsp;")) & "</a></font></DIV></DIV>" '"<li>" &  stexto1  & chr(13)
				End If

				ss7=ss7 & "</li>"
			End If
			rsx.movenext
		Loop
		Call ofv.cerrarconsulta(rsx)
		ss7=ss7 & "</UL>" & chr(13)
	End If
	armasubmNEW2=ss7
End Function

Function ARMAM()
	Dim cbmm

	Dim str
	Dim AUsrG
	str="" & Chr(13)
	'str=str & chr(13) & "<li class=lixstk TITLE='' ><a  Xclass=lixstk  href=# class='link'  > <i style=' font-size: 25px; ' class='la la-user'></i> <font size=2 color=  #fdfefe   face=verdana>USUARIO</font></a>" & Chr(13) & "</li>" & Chr(13)
	XMESAURL = ""

	If session("URLTYP") = "V" Then
		XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
	End If


	If session("URLPORT") = "N" Then
		XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
	Else
		XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
	End If



	surx=XMESAURL & "/coremanager/novedmenu.asp"

	suserinf="<i style=' font-size: 35px; ' class='la la-user'></i>  <font style=' font-size: 20px;'    face=calibri><i>" &  capitalize(Session("SegUsuario")) & "</i></font><br>"

	xx=getinfouser(Session("Usuario"))
	suserinf=suserinf & "<i style=' font-size: 18px;'>Email: " & capitalize(UCase(Session("sgu_mail")& "" ))
	suserinf=suserinf & "<BR>Grupos: "

	wktgrupos = Split(session("SegGrpdsc"), Chr(9))

	For wkti = 0 To UBound(wktgrupos)
		suserinf=suserinf & " " & capitalize(wktgrupos(wkti))
	Next

	suserinf=suserinf & "</i>"

	ARMAM="<div id=contm><div style='display:none' id=infouser>"& suserinf& "</div><div id=menucor data-cli='"&session("APLMDCLI")& "'><nav  ><h2><i class='la la-bars'></i><i class='la la-bars'></i><a class=ntmnn href='" &  surx & "' target=info><i class='la la-home'></i><font face=calibri> INICIO</font></a></h2><ul>" & chr(13)
	ARMAM=ARMAM & "<li  TITLE='' ><a class='userc ntmnn'  href=#   xstyle='background-color:#FF9B42' >  <font size=2 xcolor=  #fdfefe    face=calibri>" & Session("SegUsuario")&"&nbsp;&nbsp;<i style=' font-size: 25px; ' class='la la-user'></i></font></a>" & Chr(13) & "</li>" & Chr(13)

	If  Session("DELLOG") = "S" Then
		'  session("SegGrpLst")="MAS" & CHR(9) & "XX"
		AUsrG = Split(session("SegGrpLst"),Chr(9))
		AUsrG2 = Split(session("SegGrpLst"),Chr(9))
		XMFSEP=""
		For XMFI = 0 To UBound(AUsrG)
			UsrGS  = UsrGS  & XMFSEP & "'" & Replace(AUsrG(XMFI) & "","'","''") & "'"
			XMFSEP  = ","
		Next
	Else
		UsrGS=  "'" & Session("seggrp")  & "'"
		' AUsrG2 = Split(Session("seggrp") & chr(9) & "22222222222222222" ,Chr(9))
		AUsrG2 = Split(Session("seggrp") ,Chr(9))
		'response.write(AUsrG2(0))
	End If

	Set cmm = OFV.conectar(OFV.strconn0)
	sub1=""
	' wkaplic="EM"
	SQL= "Select max (BB.codapli) as codapli ,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr "
	SQL=SQL & " from ( "
	XMFSEP=""
	SQL2=""
	For XMFI = 0 To UBound(AUsrG2)
		UsrGS  ="'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"
		sql2=sql2 & chr(13) & XMFSEP & chr(13)
		sql2=sql2 & "Select b.codapli as codapli ,isnull(  CAST(c.grpestado AS INT)  ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1) as atrUsr"
		sql2=sql2 &  " FROM CorAplic b "
		sql2=sql2 &  " left join atrmXgrupo c on   ( c.Tipo = 'A' and c.grupo IN("& UsrGS &") and  b.codapli=C.APLICACION) "
		sql2=sql2 &  " left join atrmXusuario d on   ( D.Tipo = 'A' and d.usuario='" & Session("segperfil") & "' and D.APLICACION=b.codapli) "
		sql2=sql2 &  " where b.enable <> 0"
		XMFSEP= CHR(13) & " UNION " & CHR(13)
	Next
	SQL =SQL & SQL2 & ") BB"
	SQL=SQL & " GROUP BY BB.codapli " & CHR(13) & CHR(13)

	' sql=" Select max (b.codapli) as codapli ,MAX(isnull(  CAST(c.grpestado AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(d.usrEstado AS INT) ,1) )as atrUsr"
	' sql=  SQL & " FROM CorAplic b "
	' sql=  SQL & " left join atrmXgrupo c on   ( c.Tipo = 'A' and c.grupo IN(" & UsrGS &") and  b.codapli=C.APLICACION) "
	' sql=  SQL & " left join atrmXusuario d on   ( D.Tipo = 'A' and d.usuario='" & Session("segperfil") & "' and D.APLICACION=b.codapli) "
	' sql=  SQL & " where b.enable <> 0 "
	' sql=  SQL & " GROUP BY b.codapli"
	' session("SQL")=sql
	Set rsApp = ofv.crearconsultaEx(sql,cmm,1,parametros)
	AppH=""
	XMFSEP=""
	Do While Not rsapp.eof
		If rsAPP("atrGrupo") <> 0 And rsAPP("atrUsr") <> 0 Then
			AppH  = AppH  & XMFSEP & "'" & rsapp(0) & "'"
			XMFSEP  = ","
		End If

		rsapp.movenext
	Loop
	Call ofv.cerrarconsulta(rsapp)

	ARMAM= ARMAM &  listarM2("0",0,cmm,AppH,UsrGS)

	ARMAM= ARMAM & "</nav></div>" & sub1
	If Session("UsrCorp") And Session("EMPMUN") = "S" Then
		ARMAM= ARMAM  & chr(13) & "<div class=nnn title='Empresas' id=submEmp data-width='410' data-height='330' style=""display:none;height:280px;width:380px"">" & selecEmpre() & "</DIV>"
	End If
	ARMAM=ARMAM & "</div>"

End Function

Function selecEmpre()


	If Not Session("UsrCorp") Then
		selecEmpre="&nbsp;"
	Else

		If  Session("EMPMUN") = "S" Then

			XXXX = NEWEMPRE("NO")

			'session("ASP")="novedmenu.ASP"

			'  sEMPRE = request.querystring("EMPRE")
			'
			'  IF SEMPRE = "" THEN
			'    SEMPRE = ATREMP(0,0)
			' ELSE
			'    XXXX = NEWEMPRE(SEMPRE)
			'  END IF

			sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
			sImagen4si = "<B><FONT FACE=CorFlatIcon COLOR=#993333 SIZE=3>4</FONT></B>"
			sImagenI3 = "<B><FONT FACE=""CorFlatIcon"" COLOR=#6699CC SIZE=2>=</FONT></B>"

			SINPUT3 = "&nbsp;<a class=eee" & xsclase & " title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick=""window.top.shwDem2();"">Empresa</a>"

			'Filac    = ""
			'Celdac   = ""
			'sHref   = ""

			'sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
			'sLink   = ofv.GenLink("","",sHref,"ftemas","","",sImage,"si")
			Celdac   = ofv.GenCelda("","","","8%","","","",SINPUT3,"si")
			'Celdac   = Celdac & ofv.GenCelda(""," ","center","5%","","","",sLink,"si")
			'Celdac   = Celdac & ofv.GenCelda("","  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

			Filac    = ofv.GenRow("","","top","","","","",Celdac,"si")

			Tablac   =""' ofv.GenTabla("","aaa xstyle='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Filac,"si")

			'----------------------------------------------
			sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
			sinput2 = sinput2  & "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
			sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & " window.top.shwDem2();" & Chr(34)
			sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
			TABLAc = Tablac & "<DIV id='DEM12' STYLE='xdisplay:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:0px;left:0px;height:280px;width:380px;xborder-style:outset;xborder-color:#888888;xborder-width:2;xborder-radius: 15px 15px 15px 15px;overflow:auto;'>"


			TABLAX = TablaX & OFV.gentabla("", "aaa xstyle='border-radius: 15px 15px 15px 15px;' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
			TABLAX = TablaX & "<tr><td  "
			TABLAX = TablaX & "class=ttt colspan=3 align=left  STYLE='border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";' >"  & "&nbsp;SELECCIONE EMPRESA</td></tr>"
			TABLAX = TablaX & "<tr><td  "
			TABLAX = TablaX & " >&nbsp;</td></tr>"

			For IAEMP=1 To TAEMP

				'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

				If ATREMP(IAEMP,2) = 1 Then

					If ATREMP(IAEMP,0) = ATREMP(0,0) Then
						TABLAX = TablaX & "<tr  >"
					Else
						TABLAX = TablaX & "<tr>"
					End If

					TABLAX = TablaX & "<td width=10% "
					If ATREMP(IAEMP,0) = ATREMP(0,0) Then
						'	TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
						TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
						xbord = "" '" style='border-style:solid;border-width:1;' "
						ximg = sImagen4si & "&nbsp;"
					Else
						TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
						xbord = "  "
						ximg = sImagenI3 & "&nbsp;"
					End If




					'TABLAX = TablaX & "<td class=vti width=60% " & xbord & " >"
					TABLAX = TablaX & "<td class=vt width=60% " & xbord & " >"
					sHref = "../coress20/novedmenuEstilo.asp?EMPRE=" & ATREMP(IAEMP,0)
					'sHref = ""
					sLink = ofv.GenLink("","cc  onclick=window.top.ejec('"  & sHref & "');","#","","","","<b>" & Replace(ATREMP(IAEMP,4) & ""," ","&nbsp;") & "</b>","si")
					TABLAX = TablaX & sLink
					TABLAX = TablaX & "</td>"

					sImage = "<img src='../COREss20/Empresa/" & ATREMP(IAEMP,5) & "' width='100' >"
					'''  sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
					TABLAX = TablaX & "<td width=40% "
					TABLAX = TablaX & "class=vt  >" & sImage & "</td>"
					TABLAX = TablaX & "</tr>"
				End If

			Next

			TABLAX = TablaX & OFV.genFtabla()

			TABLAc = Tablac & TABLAX

			TABLAc = Tablac & "</DIV>"

			xsHTML = xsHTML & TABLAc

		End If
		sf="<form id=xfx name=xfx action='' method=post style='display:none;visibility:hidden;'></form>"

		selecEmpre=xsHTML & sf
	End If

End Function
Function  ListarM2(mdepen,nivel,cmm,aplic,UsrgS)
	Dim rs,rs1
	Dim rsoef



	'	sql="Select Menux.*,titapli From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & mdepen & "'"
	'	sql= sql & "ORDER BY CorAplic.CodApli"





	'	sql="Select max(a.id), a.MENU,a.mcodapli,b.titapli,max(a.menuatr) as menuatr,MAX(isnull(  CAST(c.grpestado AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(d.usrEstado AS INT) ,1) )as atrUsr, max(a.morden) as morden,max(ordenapli )  as ordenapli,max(asp )  as asp"
	'	sql= sql  &  " From Menux a "
	'	sql=sql & "inner join CorAplic b on ( b.codapli = a.mcodapli ) "
	'	sql=sql & " left join atrmXgrupo c on   (c.alias=a.menuatr and c.Tipo = 'S' and c.grupo IN(" & UsrGS &") and c.aplicacion=b.codapli) "
	'	sql =sql & " left join atrmXusuario d on   (d.alias=a.menuatr and D.Tipo = 'S' and d.usuario='" & Session("segperfil") &"' and D.aplicacion=b.codapli) "
	'      sql= sql & " where  b.CodApli in( " & aplic & " ) AND a.mdepen = '0'"
	'
	'      ' no trae formularios
	'      if session("CORMODUSOLABM") <> "S" then
	'       sql= sql & " and menuatr not in('CLIS245159671557361072846976','CLIS45312740033826860617908')"
	'      end if
	'      sql= sql & " GROUP BY b.titapli,a.mcodapli,a.MENU"
	'      sql= sql & " order by  ordenapli,b.titapli,a.mcodapli,a.MENU,morden"


	sql="Select max(bb.id), bb.MENU,bb.mcodapli,bb.titapli,max(bb.menuatr) as menuatr,MAX(isnull(  CAST(bb.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(bb.atrUsr AS INT) ,1) )as atrUsr, max(bb.morden) as morden,max(bb.ordenapli )  as ordenapli,max(bb.asp )  as asp"
	sql =sql & " from ("

	XMFSEP=""
	SQL2=""
	For XMFI = 0 To UBound(AUsrG2)
		UsrGS  ="'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"
		sql2=sql2 & chr(13) & XMFSEP & chr(13)
		sql2=sql2 &  "Select a.id, a.MENU,a.mcodapli,b.titapli,a.menuatr as menuatr,isnull(  CAST(c.grpestado AS INT)  ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1) as atrUsr, a.morden as morden,ordenapli   as ordenapli,asp  as asp "

		sql2=sql2 & "From Menux a "
		sql2=sql2 & " inner join CorAplic b on ( b.codapli = a.mcodapli ) "
		sql2=sql2 & " left join atrmXgrupo c on   (c.alias=a.menuatr and c.Tipo = 'S' and c.grupo IN("& UsrGS & ") and c.aplicacion=b.codapli) "
		sql2=sql2 & " left join atrmXusuario d on   (d.alias=a.menuatr and D.Tipo = 'S' and d.usuario='" & Session("segperfil") & "' and D.aplicacion=b.codapli) "
		sql2=sql2 & " where  b.CodApli in(" & aplic & " ) AND a.mdepen = '0'"
		' no trae formularios
		If session("CORMODUSOLABM") <> "S" Then
			sql2=sql2 & " and menuatr not in('CLIS245159671557361072846976','CLIS45312740033826860617908')"
		End If
		XMFSEP= CHR(13) & " UNION " & CHR(13)
	Next
	SQL =SQL & SQL2 & ") BB"
	sql =sql & "  GROUP BY bb.titapli,bb.mcodapli,bb.MENU"
	sql =sql & "  order by  ordenapli,bb.titapli,bb.mcodapli,bb.MENU,morden"

	'session("SQL")=sql
	'   response.write sql & "<br>"

	Set rs = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If  rs.eof Then
		rseof="S"
	Else
		rseof="N"
	End If

	Dim a(100,1) , i
	Dim imenu , iSm
	' Recorre los subdirectorios
	i=0

	If Session("UsrCorp") And Session("EMPMUN") = "S" Then
		'str=str & chr(13) & "<li class=lixstk TITLE='Empresas' ><a  Xclass=lixstk  xtitle='Empresas' href=# class='link ntmnn'  data-id='submEmp' > <i style=' font-size: 25px; ' xclass='la la-pencil'></i> <font size=2 xcolor=  #fdfefe   face=verdana><I>EMPRESAS</I></font></a>" & Chr(13) & "</li>" & Chr(13)
        str= chr(13) & "<li data-id='submEmp' class='lixstk' TITLE='EMPRESAS' ><a  class=ntmnn  xtitle='EMPRESAS' href=# class='link'  data-id='submEmp' > <i style=' font-size: 25px; ' class='la la-pencil'></i> <font size=2 xcolor=  #fdfefe   face=calibri>EMPRESAS</font></a>" & Chr(13) & "</li>" & Chr(13)
	End If
	Do While Not rs.eof
		'  ReDim Preserve a(i,0)

		If rs("atrGrupo") <> 0 And rs("atrUsr") <> 0 Then
			'if true then
			sql2="Select top 1  Menux.* From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & rs("menuatr") & "'"
			Set rs2 = ofv.crearconsultaEx(sql2,cmm,1,parametros)

			If Not rs2.eof Then
				'   response.write sql2 & chr(13)
				If rs(2)="COREM20" Then
					str=str & chr(13) & "<li data-id='subm" & rs(0) & "' class='lixstk' TITLE='" & rs("titapli") &"' ><a  class=ntmnn  xtitle='" & rs("titapli") & "' href=# class='link'  data-id='subm" & rs(0) & "' > <i style=' font-size: 25px; ' class='la la-pencil'></i> <font size=2 xcolor=  #fdfefe   face=calibri>" & Capitalize(UCase(rs("menu"))) & "</font></a>" & Chr(13) & "</li>" & Chr(13)
				Else
					str=str & chr(13) & "<li data-id='subm" & rs(0) & "'  class='lixstk' TITLE='" & rs("titapli") &"' ><a  class=ntmnn  xtitle='" & rs("titapli") & "' href=# class='link'  data-id='subm" & rs(0) & "' > <i style=' font-size: 25px; ' class='la la-gear'></i> <font size=2 xcolor=  #fdfefe   face=calibri>" & Capitalize(UCase(rs("menu"))) & "</font></a>" & Chr(13) & "</li>" & Chr(13)
				End If

				'sub menu
				'if session("SGU_userid")="CL" then

				sub1=sub1 & "<div data-color='"& session("APLITBGD") & "' class='nnn' id=subm"  & rs(0)  & " style='display:none;'> " & armasubmNEW(rs("menuatr"),cmm,rs(1),1,"'" & rs(2) & "'" ,rs(0),UsrGS) & "</DIV>"
				'else
				'sub1=sub1 & armasubm(rs("menuatr"),cmm,rs(1),1,"'" & rs(2) & "'" ,rs(0),UsrGS)
				'end if

			Else
				xsel="menuatr"
				sInput = Trim(rs(xsel))
				sHref = Trim(rs("asp"))
				sTarget = "info"
				sPath = Session("CorPath")
				If rs(3) <> "" Then sTarget = rs(3)
				spt = InStr(rs("asp"), ".asp")
				xesm = "N"
				If spt > 1 Then
					If Left(rs("asp"), spt - 1) = "Submenus" Then
						sPath = "../coress20/"
						xesm = "S"
					End If
				End If
				sseg = InStr(sHref, "?")
				If sseg = 0 Then
					sHref = sHref & "?"
				Else
					sHref = sHref & "&"
				End If
				sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")
				ssMnn = "MNN" & OFV.convertircar(UCase(rs(1)), " ", "_")
				str=str & chr(13)  & "<LI class=lixstk><font color=  #fdfefe   size=1 face=calibri ><a onclick=" & Chr(34) & "SetMnn('" & ssMnn & "');" & Chr(34)  & " class=xxxx target=info href='" & sHref & "'>" & UCase(RS("MENU")) & "</a></font></LI>"  & Chr(13)
			End If
		End If


		rs.movenext

	Loop
	Call ofv.cerrarconsulta(rs)
	str=str & "</UL>" & Chr(13)
	ListarM2=str
End Function

Function armasubm(id,cmm,idx,nivel,aplic,idx2,UsrGS)
	Dim ss7
	'dim AUsrG
	'dim rsx
	'dim XMFI

	'sql="Select Menux.*,dirapli,titapli From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & id & "' "
	SQL = "Select max(BB.id) as id, BB.MENU,BB.mcodapli,BB.titapli,max(BB.menuatr) as menuatr,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr, max(BB.morden) as morden,max(BB.ordenapli )  as ordenapli,max(BB.asp )  as asp"
	SQL= SQL & ",max(BB.target) as target,max(BB.dirapli) as dirapli"
	sql=sql  & " FROM ( " & chr(13)

	'   if  Session("DELLOG") = "S" then
	'    AUsrG = Split(session("SegGrpLst"),Chr(9))
	'   else
	'   Session("seggrp")=Session("seggrp") & chr(9) & Session("seggrp") & "222"
	'    AUsrG = Split(Session("seggrp"),Chr(9))
	'   end if

	XMFSEP=""
	sql2=""
	For XMFI = 0 To UBound(AUsrG2)
		UsrGS  =  "'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"

		sql2=sql2 & chr(13) & XMFSEP & chr(13)
		sql2=sql2 & "	Select a.id as id, a.MENU,a.mcodapli,b.titapli,a.menuatr as menuatr,isnull(  c.grpestado   ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1)as atrUsr, a.morden as morden,ordenapli   as ordenapli,asp  as asp"
		sql2=sql2 & "	, Dirapli  ,target as target"
		sql2=sql2 & "	 From Menux a "
		sql2=sql2 & "	inner join CorAplic b on ( b.codapli = a.mcodapli ) "
		sql2=sql2 & "	left join atrmXgrupo c on   ( c.alias=a.menuatr  and c.grupo IN("& UsrGS &") and c.aplicacion=b.codapli) "
		sql2=sql2 & "	left join atrmXusuario d on ( d.alias=a.menuatr  and d.usuario='"& Session("segperfil") &"' and D.aplicacion=b.codapli)"
		sql2=sql2 & "     where  b.CodApli in( " & aplic &" ) AND a.mdepen = '" & id& "'"
		'sql2= chr(13) &  XMFSEP & chr(13) & sql2
		XMFSEP  = "UNION"
	Next
	sql= sql & sql2 & chr(13) & " )BB " & chr(13)
	sql=sql & " GROUP BY BB.titapli,BB.mcodapli,BB.MENU"
	sql=sql & " order by  morden" & chr(13)  & chr(13)


	ss7=""
	'if ID="SEGURIDAD" THEN
	'session("SQL")=SQL
	'END IF

	Set rsx = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If Not rsx.eof Then
	'	If nivel=1 Then
	        if not instr(ucase(ss7),ucase("class=nnn")) then
			 ss7=ss7 & chr(13) & "<div class=nnn title='" & UCase(idx) & "'" & "id=subm"  & idx2  & " style='display:none;'><ul >"
			end if
	'	Else
		'	ss7=ss7 & chr(13) & "<div class=nnn title='" & UCase(idx) & "'" & "id=subm"  & idx2  & " style='xdisplay:none;'><ul class=ulx>"

	'	End If

		Do While Not rsx.eof

			' response.write "222"
			If rsx("atrGrupo") <> 0 And rsx("atrUsr") <> 0 Then
				STEXTO1=UCase(rsx("menu"))
				xsel="menuatr"
				sInput = Trim(rsx(xsel))
				XMESAURL = ""

				If session("URLTYP") = "V" Then
					XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
				End If


				If session("URLPORT") = "N" Then
					XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
				Else
					XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
				End If
				surx=	XMESAURL
				sHref = surx & "/" & rsx("dirapli") & "/" & Trim(rsx("asp"))
				sTarget = "info"
				sPath = rsx("dirapli")
				If rsx("target") <> "" Then sTarget = rsx("target")
				spt = InStr(rsx("asp"), ".asp")
				xesm = "N"
				If spt > 1 Then
					If Left(rsx("asp"), spt - 1) = "Submenus" Then
						sPath = "../coress20/"
						xesm = "S"
					End If
				End If
				sseg = InStr(sHref, "?")
				If sseg = 0 Then
					sHref = sHref & "?"
				Else
					sHref = sHref & "&"
				End If

				sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rsx("menu")), " ", "_")
				If instr(lcase(rsx("asp")),"base=") = 0 Then
					' agrego base=D para los menus que no son de formularios
					sHref = sHref &  "&base=D"
				End If
                sHref = sHref &  "&coraplic=" & rsx("mcodapli")
				ssMnn = "MNN" & OFV.convertircar(UCase(rsx("menu")), " ", "_")

				sAccion= sHref
				If nivel=1 Then
					If instr(lcase(sAccion),"submenus.asp") > 0 Then
						ss7 = ss7  & "<LI class=spc TITLE='" & RSX("titapli")& "'><div><font color=  #fdfefe   size=1 face=verdana ><a xtarget=info xhref='"& sAccion & "'>&nbsp;&nbsp;&nbsp;<b><u><i>" & nivel & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></i></u></b>" '"<li>" &  stexto1  & chr(13)
					Else
						ss7 = ss7  & "<LI class=spc2 TITLE='" & RSX("titapli")& "'><div><font color=  #fdfefe   size=1 face=verdana ><a xtitle='" & rsx("titapli") & "'target=info href='"& sAccion & "'><b><i>" & nivel & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></i></b>" '"<li>" &  stexto1  & chr(13)
					End If

				Else
					ss7 = ss7  & "<LI  class=lix  TITLE='" & RSX("titapli")& "' ><div><font color=  #fdfefe   size=1 face=verdana ><a xtitle='" & rsx("titapli") & "' target=info href='"& sAccion & "'>&nbsp;&nbsp;&nbsp;" & nivel &  UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font>" '"<li>" &  stexto1  & chr(13)
				End If

				ss7=ss7 & armasubm(rsx("menuatr"),cmm,rsx("menu"),nivel +1,aplic,rsx("id"),UsrGS)
				ss7=ss7 & "</div></li>"
			End If
			rsx.movenext
		Loop
		Call ofv.cerrarconsulta(rsx)
		ss7=ss7 & "</UL></div>" & chr(13)
	End If
	armasubm=ss7
End Function

Function armasubmNEW(id,cmm,idx,nivel,aplic,idx2,UsrGS)
	Dim ss7
	'dim AUsrG
	'dim rsx
	'dim XMFI

	'sql="Select Menux.*,dirapli,titapli From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & id & "' "
	SQL = "Select max(BB.id) as id, BB.MENU,BB.mcodapli,BB.titapli,max(BB.menuatr) as menuatr,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr, max(BB.morden) as morden,max(BB.ordenapli )  as ordenapli,max(BB.asp )  as asp"
	SQL= SQL & ",max(BB.target) as target,max(BB.dirapli) as dirapli"
	sql=sql  & " FROM ( " & chr(13)

	'   if  Session("DELLOG") = "S" then
	'    AUsrG = Split(session("SegGrpLst"),Chr(9))
	'   else
	'   Session("seggrp")=Session("seggrp") & chr(9) & Session("seggrp") & "222"
	'    AUsrG = Split(Session("seggrp"),Chr(9))
	'   end if

	XMFSEP=""
	sql2=""
	For XMFI = 0 To UBound(AUsrG2)
		UsrGS  =  "'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"

		sql2=sql2 & chr(13) & XMFSEP & chr(13)
		sql2=sql2 & "	Select a.id as id, a.MENU,a.mcodapli,b.titapli,a.menuatr as menuatr,isnull(  c.grpestado   ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1)as atrUsr, a.morden as morden,ordenapli   as ordenapli,asp  as asp"
		sql2=sql2 & "	, Dirapli  ,target as target"
		sql2=sql2 & "	 From Menux a "
		sql2=sql2 & "	inner join CorAplic b on ( b.codapli = a.mcodapli ) "
		sql2=sql2 & "	left join atrmXgrupo c on   ( c.alias=a.menuatr  and c.grupo IN("& UsrGS &") and c.aplicacion=b.codapli) "
		sql2=sql2 & "	left join atrmXusuario d on ( d.alias=a.menuatr  and d.usuario='"& Session("segperfil") &"' and D.aplicacion=b.codapli)"
		sql2=sql2 & "     where  b.CodApli in( " & aplic &" ) AND a.mdepen = '" & id& "'"
		'sql2= chr(13) &  XMFSEP & chr(13) & sql2
		XMFSEP  = "UNION"
	Next
	sql= sql & sql2 & chr(13) & " )BB " & chr(13)
	sql=sql & " GROUP BY BB.titapli,BB.mcodapli,BB.MENU"
	sql=sql & " order by  morden" & chr(13)  & chr(13)


	ss7=""
	'if ID="SEGURIDAD" THEN
	'session("SQL")=SQL
	'END IF

	Set rsx = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If Not rsx.eof Then
			 ss7=ss7 & chr(13) & "<ul class=spc3>"

		Do While Not rsx.eof

			' response.write "222"
			If rsx("atrGrupo") <> 0 And rsx("atrUsr") <> 0 Then
				STEXTO1=UCase(rsx("menu"))
				xsel="menuatr"
				sInput = Trim(rsx(xsel))
				XMESAURL = ""

				If session("URLTYP") = "V" Then
					XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
				End If


				If session("URLPORT") = "N" Then
					XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
				Else
					XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
				End If
				surx=	XMESAURL
				sHref = surx & "/" & rsx("dirapli") & "/" & Trim(rsx("asp"))
				sTarget = "info"
				sPath = rsx("dirapli")
				If rsx("target") <> "" Then sTarget = rsx("target")
				spt = InStr(rsx("asp"), ".asp")
				xesm = "N"
				If spt > 1 Then
					If Left(rsx("asp"), spt - 1) = "Submenus" Then
						sPath = "../coress20/"
						xesm = "S"
					End If
				End If
				sseg = InStr(sHref, "?")
				If sseg = 0 Then
					sHref = sHref & "?"
				Else
					sHref = sHref & "&"
				End If

				sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rsx("menu")), " ", "_")
				If instr(lcase(rsx("asp")),"base=") = 0 Then
					' agrego base=D para los menus que no son de formularios
					sHref = sHref &  "&base=D"
				End If
                sHref = sHref &  "&coraplic=" & rsx("mcodapli")
				ssMnn = "MNN" & OFV.convertircar(UCase(rsx("menu")), " ", "_")

				sAccion= sHref
				If nivel=1 Then
					If instr(lcase(sAccion),"submenus.asp") > 0 Then
						ss7 = ss7  & "<LI class=spc TITLE='" & RSX("titapli")& "'><div class=utx><font xclass=spc3  xcolor=  #fdfefe   xsize=1 xface=verdana ><a class=ntmnn xtarget=info xhref='"& sAccion & "'>&nbsp;&nbsp;&nbsp;<b><u><i>"  & Capitalize(Replace(STEXTO1," ","&nbsp;")) & "</a></font></i></u></b></DIV>" '"<li>" &  stexto1  & chr(13)
					    ss7=ss7 & armasubmNEW(rsx("menuatr"),cmm,rsx("menu"),nivel +1,aplic,rsx("id"),UsrGS)

					Else
						ss7 = ss7  & "<LI class=spc TITLE='" & RSX("titapli")& "'><div class=utx><font xclass=spc3  xcolor=  #fdfefe   xsize=1 xface=verdana ><a class=ntmnn xtitle='" & rsx("titapli") & "'target=info href='"& sAccion & "'><b><i>"  & Capitalize(Replace(STEXTO1," ","&nbsp;")) & "</a></font></i></b></DIV>" '"<li>" &  stexto1  & chr(13)
					End If

				Else
					ss7 = ss7  & "<LI  class=spc  TITLE='" & RSX("titapli")& "' ><div class=utx><font xclass=spc3  xcolor=  #fdfefe   xsize=1 xface=verdana ><a class=ntmnn xtitle='" & rsx("titapli") & "' target=info href='"& sAccion & "'>&nbsp;&nbsp;&nbsp;" &  Capitalize(Replace(STEXTO1," ","&nbsp;")) & "</a></font></DIV>" '"<li>" &  stexto1  & chr(13)
				End If

				ss7=ss7 & "</li>"
			End If
			rsx.movenext
		Loop
		Call ofv.cerrarconsulta(rsx)
		ss7=ss7 & "</UL>" & chr(13)
	End If
	armasubmNEW=ss7
End Function


Sub ListarM(mdepen,nivel,cmm)
	Dim rs,rs1
	Dim rsoef
	sql="Select Menux.* From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli = 'COREM20' AND Menux.mdepen = '" & mdepen & "'"
	'response.write sql & "<br>"

	Set rs = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If  rs.eof Then
		rseof="S"
	Else
		rseof="N"
	End If

	Dim a(100,1) , i
	Dim imenu , iSm
	' Recorre los subdirectorios
	i=0

	Do While Not rs.eof
		'  ReDim Preserve a(i,0)

		sql2="Select top 1 menux.id From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli = 'COREM20' AND Menux.mdepen = '" & rs("menu") & "'"
		Set rs2 = ofv.crearconsultaEx(sql2,cmm,1,parametros)

		If Not rs2.eof Then
			If nivel=0 Then
				str=str & "<li><a href='#'><i class='la la-cog'></i><font size=2 color=  #fdfefe   face=verdana>" & UCase(rs("menu")) & "</font></a><h2>" & rs("menu")   & "</h2>" & Chr(13) &"<UL>" & Chr(13)
			Else
				str=str & "<li><a href='#'><i class='la la-cog'></i><font size=2 color= #fdfefe   face=verdana>" & UCase(rs("menu")) & "</font></a><h2>" & rs("menu") & "</h2>" & Chr(13) &"<UL>" & Chr(13)
			End If

			Call ListarM(RS("MENU"),nivel + 1 ,cmm)
			str=str & "</Li>" & Chr(13)
		Else
			xsel=6
			sInput = Trim(rs(xsel))
			sHref = Trim(rs(2))
			sTarget = "info"
			sPath = Session("CorPath")
			If rs(3) <> "" Then sTarget = rs(3)
			spt = InStr(rs(2), ".asp")
			xesm = "N"
			If spt > 1 Then
				If Left(rs(2), spt - 1) = "Submenus" Then
					sPath = "../coress20/"
					xesm = "S"
				End If
			End If
			sseg = InStr(sHref, "?")
			If sseg = 0 Then
				sHref = sHref & "?"
			Else
				sHref = sHref & "&"
			End If
			sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")
			ssMnn = "MNN" & OFV.convertircar(UCase(rs(1)), " ", "_")
			str=str & "<LI class=cc><font color=  #fdfefe   size=1 face=verdana ><a onclick=" & Chr(34) & "SetMnn('" & ssMnn & "');" & Chr(34)  & " class=xxxx target=info href='" & sHref & "'>" & UCase(RS("MENU")) & "</a></font></LI>"  & Chr(13)
		End If



		rs.movenext

	Loop
	str=str & "</li></UL>" & Chr(13)
	strf=strf &  str
	str=""
End Sub

Function armasolapaComp(sCodigo,sTipo,sFuncionX)
	Dim sHTML
	Dim Celda
	bModificar = False
	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False
    If sFuncion <> "" Then sFuncion=sFuncionX
	sTipo     = request.querystring("Tipo")
	sCodigo   = request.querystring("Codigo")
	'sFuncion  = request.querystring("Funcion")
	sSClick   = request.querystring("SSClick")
	sMenu     = request.querystring("Menu")
	sSector   = request.querystring("Sector")
	sNIV1     = request.querystring("NIV1")
	sNIV2     = request.querystring("NIV2")
	sNIV3     = request.querystring("NIV3")
	sNIV4     = request.querystring("NIV4")
	sNIV5     = request.querystring("NIV5")

	sPerfil      = request.querystring("Perfil")
	sComplemento = request.querystring("Complemento")
	sComplemento = ofv.convertircar(sComplemento,"_"," ")

	''  response.write "recibe " & request.querystring & "<br>"

	wsSClick = "V" & sSClick

	If  sSClick = "" Then     sSClick      = 99


	Set cn = ofv.conectar(ofv.strconn2)



	'sHTML = "<html>"

	wxs = "<script language=JavaScript>" & Chr(13)
	wxs = wxs & " var Fbtini = 0;" & Chr(13)
	wxs = wxs & " function IniWin() { " & Chr(13)
	wxs = wxs & "  if (document.all.SSOLACT) { alert(document.SSOLACT.action);" & chr(13)
	wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
	wxs = wxs & " } else { " & chr(13)
	wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
	wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
	wxs = wxs & "  } " & chr(13)
	wxs = wxs & " } " & Chr(13)
	wxs = wxs & " </script>" & Chr(13)

	'sHTML = sHTML & wxs
   'ver onload
	'sHTML = sHTML & "<body xonload='IniWin();' >"

    sHTML=""
	Dim stabla(20)
	Dim solord(20)

	solord(0) = 0
	solord(1) = 15
	solord(2) = 1
	solord(3) = 6
	solord(4) = 3
	solord(5) = 2
	solord(6) = 4
	solord(7) = 5
	solord(8) = 7
	solord(9) = 8
	solord(10) = 11
	solord(11) = 10
	solord(12) = 9
	solord(13) = 12
	solord(14) = 13
	solord(15) = 14
	solord(16) = 16
	solord(17) = 17
	solord(18) = 18
	solord(19) = 19
	solord(20) = 20






	Dim Nombres(20)

	Dim ASPs(20)

	ASPs(0) = "imagencont2.asp?skind=1"


	If sFuncion = "E" Then
		ASPs(1) = "edGenKomp.asp?skind=1"
	Else
		ASPs(1) = "imagencontG.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(2) = "edEstrKomp.asp?skind=1"
	Else
		ASPs(2) = "imagencont3.asp?skind=1"
	End If

	ASPs(3) = "N"

	If sFuncion = "E" Then
		ASPs(4) = "edImagen2.asp?skind=1"
	Else
		ASPs(4) = "imagencont2.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(5) = "edCompuesto2.asp?skind=1"
	Else
		ASPs(5) = "ImagenIndex.asp?skind=1"
	End If




	If sFuncion = "E" Then
		ASPs(6) = "edimgarc2.asp?skind=1"
	Else
		ASPs(6) = "imagencont4.asp?skind=1"
	End If



	ASPs(7)  = "N"
	ASPs(8)  = "N"
	ASPs(9)  = "N"
	ASPs(10) = "N"
	ASPs(11) = "N"
	ASPs(12) = "N"
	ASPs(13) = "N"
	ASPs(14) = "N"
	ASPs(15) = "N"

	perfilsp = "S"







	siconmas = "ICorAcciones.jpg"



	Dim Solapas(20)
	Dim Solapas2(20)

	For z=0 To 20
		Solapas(z) = "N"
	Next



	Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


	If Session("Segco") <> "S" Then
		For xx = 1 To 20
			Nombres(xx) = "N"
		Next

		Solapas(15)  = REPLACE(RS(31)," ","*")
		Solapas2(15) = ASPs(15)
		xstrsol      = 1
		xfinal       = 1
		sSClick      = 99


	Else

		For xx = 1 To 20
			solord(xx) = 17
			Nombres(xx) = "N"
		Next



		For xx = 2 To 16
			stabla (xx - 1) = rs(xx + 30)
		Next



		yy = 0
		For xx = 1 To 15

			If cint(stabla (xx)) < 99 Then
				yy = yy + 1
				solord (stabla (xx)) = xx
			End If
		Next
		tyy = yy




		If session("ModMain") = "WK"  Then
			Solapas(0)  = "WORKFLOW"
			Solapas2(0) = ASPs(0)
			xstrsol     = 0
		Else
			xstrsol     = 1
		End If




		X=0
		For z = 2 To 16

			If rs(z) Then

				If (sFuncion = "E" And z <> 4  ) Or (sFuncion = "C" And z <> 21 And z <> 4 And z <> 61 )  Then
					x = x + 1
					Solapas(z-1)  = REPLACE(RS(z + 15)," ","*")
					Solapas2(z-1) = ASPs(z-1)




				End If
			End If

		Next


		xfinal = ubound(Solapas)

	End If

	''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

	Call ofv.cerrarconsulta(rs)
	Dim FTABELENUL
	'FTABELENUL= "<div id=menucor data-cli='"&session("APLMDCLI")& "' ><nav><h2><i class='la la-bars'></i>SOLAPAS</h2><UL>"

	FTABELENUL="<div id=contm STYLE='overflow:hidden;' xclass=sidebar><ul class=ulmenu STYLE= 'overflow:auto;Height:100%;' id=myMenu >" & chr(13)
	FTABELENUL=FTABELENUL & chr(13) & "<li><div class='item-container'><div class='item-text'><img src='" & Session("CliLogo") &"'><hr></div></div></li>" & Chr(13)

	xSolasp = ""
	xSolasp = xSolasp & "BarraCompSol.asp?Tipo=" & sTipo & "&Codigo="
	xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector
	xSolasp = xSolasp & "&Funcion=" & sFuncion
	xSolasp = xSolasp & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
	xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
	xSolasp = xSolasp & "&SSCLICK="

	r1 = ""


	If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
		xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
		xGenConta = ofv.FTitDat1

		xsclase = "xts"

	Else



		xsclase = "xts"

	End If

	xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
	xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"

	wsSClick2 = sSClick

	xsolini = 99


	For x = xstrsol To xfinal
		If solapas(solord(x)) <> "N" Then
			sSolapa =  solapas(solord(x))

			If  xsolini = 99 Then  xsolini = solord(x)

			If  sSClick = 99 Then  sSClick = xsolini

			If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


				sInput  = replace(Ucase(sSolapa),"*"," ") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"


			Else

				sInput  = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"

			End If


			If session("WEBDESK") = "S" Then


				sInput  = "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & "</font>"
				sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")
				' FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(sSolapa)  & "</a></font></LI>"

			'	FTABELENUL=FTABELENUL & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 HREF='"& xSolasp & solord(x) & "'> "  &  UCase(sSolapa) & "</a></font></LI>"
				FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(sSolapa) & "'><a class=ntmnn target=WK3 data-id='s1solapa" & solord(x) & "' class=""link"" href='"& xSolasp & solord(x) & "'>"  &  Capitalize(sSolapa)  & "</a></div></div></li>"

				Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

				If (solord(x)) = 0 Then
					wwwestilo = " style='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
				End If


				CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
				FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



				If sFuncion = "E" Or sFuncion = "ED"  Then

				Else

				End If


			Else


				t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
				t3 = t3 & sInput
				t3 = t3 & "</A>"



				t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

				r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

			End If

			If clng(sSClick) = clng(solord(x))  and false Then

				xverasp = solapas2(solord(x)) & "&Tipocomp=" & sTipo & "&Codcomp=" & sCodigo & "&id=" & sCodigo
				xverasp = xverasp & "&Menu=" & sMenu & "&Sector=" & sSector
				xverasp = xverasp & "&GLBSOLN=S" & solord(x)
				xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

				sHTML1 = "" & chr(13)
				sHTML1 = sHTML1 & "<form name=SSOLACT action='" & xverasp & "' "
				sHTML1 = sHTML1 & " method=post target='Principal' style='display:none;' >" & chr(13)
				sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
				sHTML1 = sHTML1 & "</form>" & chr(13)

				sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
				sHTML1 = sHTML1 & " window.top.CorSol='" & replace(Ucase(sSolapa),"*"," ") & "'; " & Chr(13)
				sHTML1 = sHTML1 & "</script>" & Chr(13)


				If session("WEBDESK") = "S" Then




				Else

					If  Session("FULLGUI") = "N" Then
						styl = "background-color:#888888;"
					Else
						styl = XspEx("#888888","#cccccc",0) & ";"
					End If

					xlfun = " onclick=""document.all.MDBEDL.style.display='block'; document.all.DBEDTD1.style.display='none'; document.all.DBEDTD2.style.display='block'; document.all.MM2.style.display='block'; "" "
					xlfun2 = " onclick=""document.all.MDBEDL.style.display='none'; document.all.DBEDTD1.style.display='block'; document.all.DBEDTD2.style.display='none'; document.all.MM2.style.display='none'; "" "

					macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

					xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


					macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
					macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

					marow = ofv.GenRow("","","","","20","","",macel,"si")

					'sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

				End If

			End If


		End If
	Next

	If session("WEBDESK") = "S" Then


		'''    sHTML = sHTML & Celda
		'''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

		'sHTML = sHTML & "<div id='WKTBSP1'  >"
	'	sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "aaa  ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN, "si")
		'sHTML = sHTML & "</div>"

	Else


		'sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"

		If  Session("FULLGUI") = "N" Then
			styl = "background-color:#888888;"
		Else
			styl = XspEx("#888888","#cccccc",0) & ";"
		End If

		styl = ""

		If ofv.ModTitCls = "1" Then

		'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

			''    sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,-2,-2,"","",r1,"si")

		Else

		'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			''   sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,2,2,"","",r1,"si")

		End If

	'	sHTML = sHTML & "</div>"

	End If
	FTABELENUL=FTABELENUL &"</UL></DIV>"
	armasolapaComp= FTABELENUL

	Call ofv.cerrarconn(cn)

	'sHTML = sHTML & "</body></html>"





End Function


Function armasolapaComp2ww(sCodigo,sTipo,sFuncionX)
	Dim sHTML
	Dim Celda
	bModificar = False
	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False

	sTipo     = request.querystring("Tipo")
	sCodigo   = request.querystring("Codigo")
	sFuncion  = request.querystring("Funcion")
	sSClick   = request.querystring("SSClick")
	sMenu     = request.querystring("Menu")
	sSector   = request.querystring("Sector")
	sNIV1     = request.querystring("NIV1")
	sNIV2     = request.querystring("NIV2")
	sNIV3     = request.querystring("NIV3")
	sNIV4     = request.querystring("NIV4")
	sNIV5     = request.querystring("NIV5")

	sPerfil      = request.querystring("Perfil")
	sComplemento = request.querystring("Complemento")
	sComplemento = ofv.convertircar(sComplemento,"_"," ")

	''  response.write "recibe " & request.querystring & "<br>"

	wsSClick = "V" & sSClick

	'If  sSClick = "" Then     sSClick      = 99


	Set cn = ofv.conectar(ofv.strconn2)



	sHTML = "<html>"

	wxs = "<script language=JavaScript>" & Chr(13)
	wxs = wxs & " var Fbtini = 0;" & Chr(13)
	wxs = wxs & " function IniWin() { " & Chr(13)
	wxs = wxs & "  if (document.all.SSOLACT) { " & chr(13)
	wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
	wxs = wxs & " } else { " & chr(13)
	wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
	wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
	wxs = wxs & "  } " & chr(13)
	wxs = wxs & " } " & Chr(13)
	wxs = wxs & " </script>" & Chr(13)

	sHTML = sHTML & wxs

	sHTML = sHTML & "<body onload='IniWin();' >"


	Dim stabla(20)
	Dim solord(20)

	solord(0) = 0
	solord(1) = 15
	solord(2) = 1
	solord(3) = 6
	solord(4) = 3
	solord(5) = 2
	solord(6) = 4
	solord(7) = 5
	solord(8) = 7
	solord(9) = 8
	solord(10) = 11
	solord(11) = 10
	solord(12) = 9
	solord(13) = 12
	solord(14) = 13
	solord(15) = 14
	solord(16) = 16
	solord(17) = 17
	solord(18) = 18
	solord(19) = 19
	solord(20) = 20






	Dim Nombres(20)

	Dim ASPs(20)

	ASPs(0) = "imagencont2.asp?skind=1"


	If sFuncion = "E" Then
		ASPs(1) = "edGenKomp.asp?skind=1"
	Else
		ASPs(1) = "imagencontG.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(2) = "edEstrKomp.asp?skind=1"
	Else
		ASPs(2) = "imagencont3.asp?skind=1"
	End If

	ASPs(3) = "N"

	If sFuncion = "E" Then
		ASPs(4) = "edImagen2.asp?skind=1"
	Else
		ASPs(4) = "imagencont2.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(5) = "edCompuesto2.asp?skind=1"
	Else
		ASPs(5) = "ImagenIndex.asp?skind=1"
	End If




	If sFuncion = "E" Then
		ASPs(6) = "edimgarc2.asp?skind=1"
	Else
		ASPs(6) = "imagencont4.asp?skind=1"
	End If



	ASPs(7)  = "N"
	ASPs(8)  = "N"
	ASPs(9)  = "N"
	ASPs(10) = "N"
	ASPs(11) = "N"
	ASPs(12) = "N"
	ASPs(13) = "N"
	ASPs(14) = "N"
	ASPs(15) = "N"

	perfilsp = "S"







	siconmas = "ICorAcciones.jpg"



	Dim Solapas(20)
	Dim Solapas2(20)

	For z=0 To 20
		Solapas(z) = "N"
	Next



	Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


	If Session("Segco") <> "S" Then
		For xx = 1 To 20
			Nombres(xx) = "N"
		Next

		Solapas(15)  = REPLACE(RS(31)," ","*")
		Solapas2(15) = ASPs(15)
		xstrsol      = 1
		xfinal       = 1
		sSClick      = 99


	Else

		For xx = 1 To 20
			solord(xx) = 17
			Nombres(xx) = "N"
		Next



		For xx = 2 To 16
			stabla (xx - 1) = rs(xx + 30)
		Next



		yy = 0
		For xx = 1 To 15

			If cint(stabla (xx)) < 99 Then
				yy = yy + 1
				solord (stabla (xx)) = xx
			End If
		Next
		tyy = yy




		If session("ModMain") = "WK"  Then
			Solapas(0)  = "WORKFLOW"
			Solapas2(0) = ASPs(0)
			xstrsol     = 0
		Else
			xstrsol     = 1
		End If




		X=0
		For z = 2 To 16

			If rs(z) Then

				If (sFuncion = "E" And z <> 4  ) Or (sFuncion = "C" And z <> 21 And z <> 4 And z <> 61 )  Then
					x = x + 1
					Solapas(z-1)  = REPLACE(RS(z + 15)," ","*")
					Solapas2(z-1) = ASPs(z-1)




				End If
			End If

		Next


		xfinal = ubound(Solapas)

	End If

	''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

	Call ofv.cerrarconsulta(rs)
	Dim FTABELENUL
	FTABELENUL= "<div style='display:none' id=menucor data-cli='"&session("APLMDCLI")& "' ><nav><h2><i class='la la-bars'></i>SOLAPAS</h2><UL>"
	xSolasp = ""
	xSolasp = xSolasp & "BarraCompSol.asp?Tipo=" & sTipo & "&Codigo="
	xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector
	xSolasp = xSolasp & "&Funcion=" & sFuncion
	xSolasp = xSolasp & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
	xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
	xSolasp = xSolasp & "&SSCLICK="

	r1 = ""


	If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
		xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
		xGenConta = ofv.FTitDat1

		xsclase = "xts"

	Else



		xsclase = "xts"

	End If

	xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
	xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"

	wsSClick2 = sSClick

	xsolini = 99


	For x = xstrsol To xfinal
		If solapas(solord(x)) <> "N" Then
			sSolapa =  solapas(solord(x))

			If  xsolini = 99 Then  xsolini = solord(x)

			If  sSClick = 99 Then  sSClick = xsolini

			If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


				sInput  = replace(Ucase(sSolapa),"*"," ") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"


			Else

				sInput  = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"

			End If


			If session("WEBDESK") = "S" Then


				sInput  = "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & "</font>"
				sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")
				' FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(sSolapa)  & "</a></font></LI>"
				FTABELENUL=FTABELENUL & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 HREF='"& xSolasp & solord(x) & "'> "  &  UCase(sSolapa) & "</a></font></LI>"
				Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

				If (solord(x)) = 0 Then
					wwwestilo = " style='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
				End If


				CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
				FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



				If sFuncion = "E" Or sFuncion = "ED"  Then

				Else

				End If


			Else


				t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
				t3 = t3 & sInput
				t3 = t3 & "</A>"



				t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

				r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

			End If

			If clng(sSClick) = clng(solord(x)) Then

				xverasp = solapas2(solord(x)) & "&Tipocomp=" & sTipo & "&Codcomp=" & sCodigo & "&id=" & sCodigo
				xverasp = xverasp & "&Menu=" & sMenu & "&Sector=" & sSector
				xverasp = xverasp & "&GLBSOLN=S" & solord(x)
				xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

				sHTML1 = "" & chr(13)
				sHTML1 = sHTML1 & "<form name=SSOLACT action='" & xverasp & "' "
				sHTML1 = sHTML1 & " method=post target='Principal' style='display:none;' >" & chr(13)
				sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
				sHTML1 = sHTML1 & "</form>" & chr(13)

				sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
				sHTML1 = sHTML1 & " window.top.CorSol='" & replace(Ucase(sSolapa),"*"," ") & "'; " & Chr(13)
				sHTML1 = sHTML1 & "</script>" & Chr(13)


				If session("WEBDESK") = "S" Then




				Else

					If  Session("FULLGUI") = "N" Then
						styl = "background-color:#888888;"
					Else
						styl = XspEx("#888888","#cccccc",0) & ";"
					End If

					xlfun = " onclick=""document.all.MDBEDL.style.display='block'; document.all.DBEDTD1.style.display='none'; document.all.DBEDTD2.style.display='block'; document.all.MM2.style.display='block'; "" "
					xlfun2 = " onclick=""document.all.MDBEDL.style.display='none'; document.all.DBEDTD1.style.display='block'; document.all.DBEDTD2.style.display='none'; document.all.MM2.style.display='none'; "" "

					macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

					xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


					macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
					macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

					marow = ofv.GenRow("","","","","20","","",macel,"si")

					sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

				End If

			End If


		End If
	Next

	If session("WEBDESK") = "S" Then


		'''    sHTML = sHTML & Celda
		'''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

		sHTML = sHTML & "<div id='WKTBSP1'  >"
		sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "aaa  ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN, "si")
		sHTML = sHTML & "</div>"

	Else


		sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"

		If  Session("FULLGUI") = "N" Then
			styl = "background-color:#888888;"
		Else
			styl = XspEx("#888888","#cccccc",0) & ";"
		End If

		styl = ""

		If ofv.ModTitCls = "1" Then

			sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

			''    sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,-2,-2,"","",r1,"si")

		Else

			sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			''   sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,2,2,"","",r1,"si")

		End If

		sHTML = sHTML & "</div>"

	End If
	FTABELENUL=FTABELENUL &"</UL></NAV></DIV>"
	armasolapaComp2= FTABELENUL

	Call ofv.cerrarconn(cn)

	'sHTML = sHTML & "</body></html>"





End Function

Function armasolapa2ww(sCodigo,sTipo,sFuncionX)

	Dim sHTML
	Dim Celda
	ok=ofv.CheckUsuario()
	If ok <> "Y" Then
		uv = ofv.uv
		If uv = "NO" Then ok = ofv.ValidarUsuario
		sHTML = ok
	Else
		'response.expires=0
		bModificar = False
		ofv.ObtenerAtributos Session("Form2"),""
		bModificar = ofv.AModificar
		sCodigo=session("ncliente")
		sTipo=session("tcliente")

		' ver ale 9/4/2021
		' se toma por request.querystring ahora
		'if sCodigo="" then
		sTipo     = request.querystring("Tipo")
		sCodigo   = request.querystring("Codigo")
		' end if

		If session("APLMDCLI")="" Then
			aplcheck()
		End If

		If session("APLMDCLI")="BGAL" Then
			lengtd=""
		Else
			lengtd="20"
		End If
		iitemC=0
		If session("Nomodificar") = "S" Then bModificar = False

		Select Case session("ModMain")

			Case "ED"
			sFuncion="E"
			Case "CL"
			sFuncion="C"
			Case "WK"
			sFuncion="CD"



		End Select

		If sFuncion <> "" Then sFuncion=sFuncionX

		sSClick   = request.querystring("SSClick")
		sMenu     = request.querystring("Menu")
		sSector   = request.querystring("Sector")
		sNIV1     = request.querystring("NIV1")
		sNIV2     = request.querystring("NIV2")
		sNIV3     = request.querystring("NIV3")
		sNIV4     = request.querystring("NIV4")
		sNIV5     = request.querystring("NIV5")
		sMAINAPL  = request.querystring("MAINAPL")

		''Nuevo para Test
		sTest     = request.querystring("Test")

		' response.write sNIV1 & " &Test "

		''Nuevo para Test

		'response.write "recibe " & session("ModMain") & "<br>"

		Set cn = ofv.conectar(ofv.strconn2)
		Strsql = "Select TIPOCPT From TIPOBJETO Where TIPOBJETO ='" & sTipo & "'"
		Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		If Not rsx1.eof Then
			TIPOCPT=rsx1(0)
		Else
			TIPOCPT="D"
		End If
		Call ofv.cerrarconsulta(rsx1)


		'sHTML = "<html>"
		'sHTML = sHTML & "<HEAD><META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " content='text/html; charset=" & XLCHARSET & "'></HEAD>"

		wxs = "<script language=JavaScript>" & Chr(13)
		wxs = wxs & " var Fbtini = 0;" & Chr(13)
		wxs = wxs & " function IniWin() { " & Chr(13)
		wxs = wxs & "  if (document.getElementById('SSOLACT')) { " & Chr(13)
		wxs = wxs & " document.getElementById('SSOLACT').submit(); " & Chr(13)
		wxs = wxs & " } else { " & Chr(13)
		wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & Chr(13)
		wxs = wxs & " setTimeout (""IniWin()"", 100); } " & Chr(13)
		wxs = wxs & "  } " & Chr(13)
		wxs = wxs & " } " & Chr(13)
		wxs = wxs & " </script>" & Chr(13)

		'sHTML = sHTML & wxs

		'sHTML = sHTML & "<body onload='IniWin();' >"



		If sFuncion = "E" Then
			smodosol = "SP"
		Else
			smodosol = "TB"
		End If


		Dim stabla(20)
		Dim solord(20)
		Dim Nombres(20)
		Dim ASPs(20)
		Dim IKNs(20)
		Dim Solapas(20)
		Dim Solapas2(20)
		Dim SolapasI(20)
		Dim rs

		ASPs(0) = "CoWKFselect.asp?SKIND=1"

		wwwsaltador = "<BR>"


		If SMAINAPL = "OR" Or SMAINAPL = "PL" Then



			ASPs(0) = "COWKFVERIFICACION.asp?SKIND=1"

			solord(0) = 0
			solord(1) = 1
			solord(2) = 2
			solord(3) = 3
			solord(4) = 4
			solord(5) = 5
			solord(6) = 6
			solord(7) = 7
			solord(8) = 8
			solord(9) = 9
			solord(10) = 10
			solord(11) = 11
			solord(12) = 12
			solord(13) = 13
			solord(14) = 14
			solord(15) = 15
			solord(16) = 16
			solord(17) = 17
			solord(18) = 18
			solord(19) = 19
			solord(20) = 20



			perfilsp = "S"

			If SMAINAPL = "OR" Then
				sEsquema = "CIRCUITO AUTOEVALUACION"
			Else
				sEsquema = "PLAN DE CONTINGENCIAS"
			End If

			xestado = 1

			Strsql = "Select estado,CODOBJETO,TIPOBJETO,RELLINK From ORRSKCIRCUITO Where OBJID = 0" & sCodigo
			Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsx1.eof Then
				xestado = rsx1(0)
				sCodigoVINC = rsx1(1)
				sTIPOVINC = rsx1(2)
				sRELVINC = rsx1(3)
			End If
			Call ofv.cerrarconsulta(rsx1)


			If (sFuncion = "CD") Then
				ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=0"
			Else
				ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=1"
			End If




			For z=0 To 20
				Solapas(z) = "N"
			Next


			For xx = 1 To 20
				Nombres(xx) = "N"
			Next

			'Solapas(1)  = REPLACE("DETALLE DEL CIRCUITO"," ","*")
			Solapas(1)  = Replace("MATRIZ DE AUTOEVALUACION"," ","*")
			Solapas2(1) = ASPs(1)
			xstrsol      = 1
			xfinal       = 1
			'''   sSClick      = 99



			If session("ModMain") = "WK"  Then
				Solapas(0)  = "WORKFLOW"
				Solapas2(0) = ASPs(0)
				SolapasI(0) = "#e7ea1c"
				xstrsol     = 0

			Else
				xstrsol     = 1
			End If



			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		Else


			ASPs(1) = "edPropiedades.asp?SKIND=1"
			IKNs(1) = "#baffc9"

			If sFuncion = "E" Then
				''' ASPs(2) = "edConsideraciones.asp?SKIND=1"
				ASPs(2) = "EDNWEDT.asp?SKIND=1"
			Else
				'''    ASPs(2) = "CoConsideraciones.asp?SKIND=1"
				ASPs(2) = "CONWEDT.asp?SKIND=1"
			End If

			IKNs(2) = "#ffffba"

			If sFuncion = "E" Then
				ASPs(3) = "edSectores.asp?SKIND=1"
			Else
				ASPs(3) = "CoSectores.asp?SKIND=1"
			End If

			IKNs(3) = "#bae1ff"



			ASPs(4) = "edEventos.asp?SKIND=1"

			IKNs(4) = "#ff7f50"

			If sFuncion = "E" Then
				ASPs(5) = "edProcedimientos.asp?SKIND=1"
			Else
				ASPs(5) = "CoProcedimientos.asp?SKIND=1"
			End If

			IKNs(5) = "#ffdfba"

			If sFuncion = "E" Then
				ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
			Else
				ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
			End If

			IKNs(16) = "#ffb3ba"

			If sFuncion = "E" Then
				ASPs(6) = "edObjDsc.asp?SKIND=1"
			Else
				ASPs(6) = "CoObjDsc.asp?SKIND=1"
			End If

			IKNs(6) = "#3dece2"

			If sFuncion = "E" Then
				ASPs(7) = "edControles.asp?SKIND=1"
			Else
				ASPs(7) = "CoControles.asp?SKIND=1"
			End If

			IKNs(7) = "#e397e2"

			If sFuncion = "E" Then
				ASPs(8) = "edEsquemas.asp?SKIND=1"
			Else
				ASPs(8) = "CoEsquemas.asp?SKIND=1"
			End If

			IKNs(8) = "#f6d2d2"

			If sFuncion = "E" Then
				ASPs(9) = "edTablas.asp?SKIND=1"
			Else
				ASPs(9) = "CoTablasCons.asp?SKIND=1"
			End If

			IKNs(9) = "#97e3c6"

			If sFuncion = "E" Then
				ASPs(10) = "edarchivos.asp?SKIND=1"
			Else
				ASPs(10) = "CoArchCons.asp?SKIND=1"
			End If

			IKNs(10) = "#ddc37a"

			If sFuncion = "E" Then
				ASPs(11) = "edestr.asp?SKIND=1"
			Else
				''  ASPs(11) = "Coestrex.asp?SKIND=1"
				ASPs(11) = "CoestrCons.asp?SKIND=1"
			End If

			IKNs(11) = "#c0db50"

			If sFuncion = "E" Then
				ASPs(12) = "CursogramaLnch.asp?SKIND=1"
				'''   ASPs(12) = "edProcedimientos.asp?SKIND=1"
			Else
				ASPs(12) = "CursogramaLnch.asp?SKIND=1"
				''''' ASPs(12)="CursogramaView.asp?SKIND=2"
			End If

			IKNs(12) = "#66cdaa"

			sMAINAPL = request.querystring("MAINAPL")
			If sFuncion = "E" Then
				ASPs(13) = "edarchivos.asp?SKIND=2"
			Else
				ASPs(13) = "CoArchCons.asp?SKIND=2"
			End If

			''''  IKNs(13) = "#ddffc9"


			If sFuncion = "E" Then
				ASPs(14) = "edFAQ.asp?SKIND=1"
			Else
				ASPs(14) = "CoFAQ.asp?SKIND=1"
			End If

			'---- IKNs(14) = "#ddffc9"


			If sFuncion = "E" Then


				ASPs(15) = "edGeneral.asp?SKIND=1"
			Else


				If TIPOCPT="K" Then
					ASPs(15) = "CoGeneralForm.asp?SKIND=1"
				Else
					ASPs(15) = "CoGeneral.asp?SKIND=1"
				End If
				' ASPs(15) = "CoGeneral.asp?SKIND=1"
			End If

			IKNs(15) = "#ddffc9"

			If sFuncion = "E" Then
				ASPs(20) = "edportada.asp?SKIND=1"
			Else
				ASPs(20) = "Coportada.asp?SKIND=1"
			End If

			IKNs(20) = "#f6d2d2"

			solord(0) = 0
			solord(1) = 15
			solord(2) = 1
			solord(3) = 6
			solord(4) = 3
			solord(5) = 2
			solord(6) = 4
			solord(7) = 5
			solord(8) = 7
			solord(9) = 8
			solord(10) = 11
			solord(11) = 10
			solord(12) = 9
			solord(13) = 12
			solord(14) = 13
			solord(15) = 14
			solord(16) = 16
			solord(17) = 17
			solord(18) = 18
			solord(19) = 19
			solord(20) = 20




			perfilsp = "S"


			siconmas = "ICorAcciones.jpg"
			Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
			'	response.write strsql & "<br>"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs.eof Then
				sEsquema = rs(1)
				sPerfil = rs(3)
				COMENTARIOS=RS("COMENTARIOS")
			End If

			Call ofv.cerrarconsulta(rs)


			For z=0 To 20
				Solapas(z) = "N"
			Next




			Strsql = "Select * From Objperfildoc Where codobjeto =0" & sCodigo & " and tipobjeto ='" & stipo & "' "
			'	  response.write strsql
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If rs.eof Then
				perfilsp = "N"
			End If

			If perfilsp = "N" Then
				Call ofv.cerrarconsulta(rs)
				Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"

				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			End If

			If Session("Segco") <> "S" Then
				For xx = 1 To 20
					Nombres(xx) = "N"
				Next

				Solapas(15)  = Replace(RS(31)," ","*")
				Solapas2(15) = ASPs(15)
				SolapasI(15) = IKNs(15)
				xstrsol      = 1
				xfinal       = 1
				sSClick      = 99


			Else

				For xx = 1 To 20
					solord(xx) = 17
					Nombres(xx) = "N"
				Next



				For xx = 2 To 16
					stabla (xx - 1) = rs(xx + 30)
				Next



				yy = 0
				For xx = 1 To 15

					If CInt(stabla (xx)) < 99 Then
						yy = yy + 1
						'   response.write stabla (xx) & "=" &  xx & chr(13)
						solord (stabla (xx)) = xx
					End If
				Next
				tyy = yy



				'''  if session("ModMain") = "WK" and sFuncion = "E" then
				If session("ModMain") = "WK"  Then
					'  if session("ModMain") = "WK" and ( session("MODSTART") <> "XP" or (session("MODSTART") = "XP" and session("WKFCODIGO") = sCodigo) ) then
					Solapas(0)  = "WORKFLOW"
					Solapas2(0) = ASPs(0)
					SolapasI(0) = "#e7ea1c"
					''' xstrsol     = 0
					xstrsol     = 0

					If session("APLMDCLI") <> "BC" Then


						' se saco porque generaba error con la solapa propiedades en el primer lugar 22_9_2017

						'       if sSClick      = 99 then sSClick      = solord(1)

					End If
				Else
					xstrsol     = 1
				End If





				X=0

				If  sFuncion <> "E"  Then

					'if sFuncion <> "E" then

					''' Solapas(20)  = "PORTADA"
					Solapas2(20) = ASPs(20)
					SolapasI(20) = IKNs(20)

					Strsql = "SELECT OBJportada.detcon, OBJportada.detsecuen, OBJportada.dettipo from OBJportada "
					Strsql = Strsql & " where OBJportada.codobjeto=" & sCodigo & " and OBJportada.detcon > 0 and OBJportada.dettipo <> 'N' "

					' response.write strsql
					Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If Not rsxx1.eof Then

					End If

					For xiord=0 To 19

						If solord(xiord) = 1 Then

							xmovey = 20

							'  for xiord2=xiord + 1 to 19
							'      xmoven           = solord(xiord2)
							'      solord(xiord2)   = xmovey
							'      xmovey           = xmoven
							'  next



							Exit For
						End If
					Next


					WWWMOVER = 20

					For xiord=1 To 20

						XXXMOVER      = solord(xiord)
						solord(xiord) = WWWMOVER
						WWWMOVER      = XXXMOVER

					Next




				End If

				For z = 2 To 16

					If rs(z) Then




						If (sFuncion = "E"  ) Or (sFuncion = "C" And z <> 2 And z <> 4 And z <> 5 ) Or (sFuncion = "CD" And z <> 2 And z <> 5 )  Or (sFuncion = "D"  And z = 4 ) Then
							x = x + 1
							Solapas(z-1)  = Replace(RS(z + 15)," ","*")
							Solapas2(z-1) = ASPs(z-1)
							SolapasI(z-1) = IKNs(z-1)

							If sFuncion <> "E" And z = 6 Then

								Solapas(16)  = "CIRCUITO COMPLETO"
								Solapas2(16) = ASPs(16)
								SolapasI(16) = IKNs(16)

								For xiord=0 To 19

									If solord(xiord) = 5 Then

										xmovey = 16

										For xiord2=xiord + 1 To 19
											xmoven           = solord(xiord2)
											solord(xiord2)   = xmovey
											xmovey           = xmoven
										Next



										Exit For
									End If
								Next


							End If

						End If
					End If

				Next


				xfinal = UBound(Solapas)

				' response.write "sSClick=" & sSClick & " solord(1) " &  solord(1)& chr(13)
				'  if sSClick      = 1 then sSClick      = solord(1)
			End If

			''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

			Call ofv.cerrarconsulta(rs)


		End If


		TABELEN = ""
		FTABELEN = ""
		Dim FTABELENUL
		Dim ars
		FTABELENUL= "<div id=menucor data-cli='"&session("APLMDCLI")& "'><nav><h2><i class='la la-bars'></i>SOLAPAS</h2><UL>"



		xSolasp = ""
		xSolasp = xSolasp & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
		xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
		xSolasp = xSolasp & "&Funcion=" & sFuncion
		xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
		xSolasp = xSolasp & "&SSCLICK="


		xSolaspd = ""
		xSolaspd = xSolaspd & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
		xSolaspd = xSolaspd & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
		xSolaspd = xSolaspd & "&Funcion=" & sFuncion
		xSolaspd = xSolaspd & "&SSCLICK="

		r1 = ""


		If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
			xGenConta = ofv.GetTit(CLng(ofv.ModTitCls) + 5)
			xGenConta = ofv.FTitDat1

			xsclase = "xts"

		Else



			xsclase = "xts"

		End If

		xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
		xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"



		xsolini = 99


		If session("WEBDESK") = "S" And session("ModMain") = "WK" Then

			sInput  = "<font color='#990000' >"
			sInput  = sInput & "Este documento requiere su confirmacion en el workflow!!!</font>"
			sLink   = sInput

			CTABELEN = ofv.GenCelda("","vT  style='font-size:8pt;background-color:#ffffff;border-style:solid;border-width:1;border-color:#000000;padding-left:5px;' ","","","","","",sLink,"si")
			'  FTABELEN = FTABELEN & ofv.GenRow("","","","","25","","",CTABELEN,"si")

		End If

		If session("WEBDESK") = "S" And session("ModMain") <> "WK" And (sFuncion = "C" Or sFuncion = "CD") Then

			xverasp = "CoObjRecomendar.asp?Tipobjeto=" & sTipo & "&Codobjeto="
			xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
			xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5


			sInput  = "<b><font color='#ff0000' size=2  >&#058;&#0183;)</font></b>&nbsp;&nbsp;"
			'''  sInput  = ""
			sInput  = sInput & "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >Recomendar </font>"
			'''  sInput  = sInput & "&nbsp;<font color='#ff0000' size=2 >&#033;</font>"
			sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:bold;'  ", xverasp, "Principal", "", "", sInput, "si")
			Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

			wwwestilo = " style='background-color:#FFEEDD;border-style:dotted;border-width:2;border-color:#990000;padding-left:10px;' "

			CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
			' FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")

		End If

		wwwk1 = "<div id='portadaFR' name='contenedor' class='portada' style='display:none;' >"

		'response.write   xstrsol & " to  "  & xfinal &  chr(13)
		For x = xstrsol To xfinal
			'' for x = xstrsol to 2
			'response.write solapas(solord(x)) & " x=" & x &  chr(13)
			'response.write  solapas(solord(x)) & "<br>"
			If solapas(solord(x)) <> "N" Then
				sSolapa =  solapas(solord(x))
				sikon   =  SolapasI(solord(x))


				If  xsolini = 99 Then  xsolini = solord(x)

				If  sSClick = 99 Then  sSClick = xsolini

				If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


					If sFuncion = "E" Or sFuncion = "ED"  Then
						sInput  = Replace(UCase(sSolapa),"*"," ")
					Else
						sInput  = "<img src='" & sikon & "' >"

					End If




					xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:10px;'"


				Else


					If sFuncion = "E" Or sFuncion = "ED"  Then
						sInput  = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
					Else
						'''  sInput  = "<img width='80px' src='" & sikon & "' >"
						sInput  = "<input type='button' value="'" & replace(Ucase(sSolapa),"*"," ") & "'" style='font-weight:bold; font-size:16px; background-color: #baffc9; border-radius: 60px; width: 167px; height: 40px'>"
					End If




					xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:10px;'"

				End If

				If session("WEBDESK") = "S" Then


					Dim Nombre_Solapa
					If sFuncion = "E" Or sFuncion = "ED"  Then
						sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >"
						sInput  = sinput & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;")
						sInput  = sinput & "</font>"
						Nombre_Solapa = Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," "," ")
					Else
						'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >"
						'sInput  = sinput & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;")
						'sInput  = sinput & "</font>"
						'sInput  = sinput & "<br>"
						sInput = ""
						''  sInput  = sinput & "<img width='80%;' height='50px;' src='" & sikon & "'  title='" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & "' "
						''  sInput  = sinput & " xonmouseover=""alert('entra SOLX" & solord(x) & "');"" xonmouseout=""alert('sale SOLX" & solord(x) & "');"" >"

						'sInput  = "<input type='button' value='" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," "," ") & "' style='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 40px; border-style:solid;border-width:1px;border-color:grey;'>"

						'sInput  = "<input type='button' value='" & left(replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," "," "),20) & vbCrLf & mid(replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," "," "),21,40) & "' style='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 50px; border-style:solid;border-width:1px;border-color:grey;'>"


						Nombre_Solapa = Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," "," ")
						'Dim Nombre_Solapa1
						'Nombre_Solapa1 =  left(replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," "," "),10)
						'Dim Nombre_Solapa2
						'Nombre_Solapa2 =  mid(replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," "," "),11,40)
						'Dim Nombre_Solapa3
						'Nombre_Solapa3 =  replace(Nombre_Solapa2,"nbsp;","<br>")

						'sInput  = "<div ' style='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 50px; border-style:solid;border-width:1px;border-color:grey;'>" & left(replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," "," "),20) & "<br>" & mid(replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," "," "),21,40) & "</div>"

						'sInput  = "<div ' style='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 50px; border-style:solid;border-width:1px;border-color:grey;'>" & Nombre_Solapa1 & Nombre_Solapa2 & "</div>"

						sInput  = "<div id='LNK" & solord(x) & "'  valign= 'middle' style='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 50px; border-style:solid; border-width:1px; border-color:grey; display:flex; justify-content:center; align-content:center; flex-direction:column;'>" & Nombre_Solapa & "</div>"

						'sInput  = replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;")

						''''sInput  = sinput & " onmouseOver=""document.getElementById('SOLX" & solord(x) & "').style.display:block;"" onmouseout=""document.getElementById('SOLX" & solord(x) & "').style.display:none;"" >"
					End If


					'   ssMnnSOL = "MNNSOL" & X
					iitemC=iitemC+1
					'   sLink   = "<span id='" & ssMnnSOL & "'  class='XX' >" & sInput & "</span> "
					'   sLink   = OFV.GenLink("", "cmb  onclick=" & chr(34) & "SetMnnSOL('"&ssMnnSOL&"');" & chr(34) & "style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")


					If sFuncion = "E" Or sFuncion = "ED"  Then
						sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput , "si")

					Else
						sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput , "si")
					End If

					'     sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput , "si")
					'
					Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink  & "&nbsp;&nbsp;</span> "

					wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

					If (solord(x)) = 0 Then
						wwwestilo = " style='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
					End If

					CTABELEN = ofv.GenCelda("","uT align='center' " & wwwestilo & " ","","","","","",sLink,"si")
					FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

					SLI=""
					Select Case solord(x)



						Case 5

						eveant = ""

						Strsql = "SELECT *  "
						Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & sCodigo
						strsql = strsql & " order by secuen, secuencia,codproced "
						'  RESPONSE.WRITE STRSQL
						sInput=""

						'FTABELENUL= FTABELENUL & "<li class=solapa" & solord(x) & " ><a class=solapa" & solord(x) & " href='#'><i class='la la-cog'></i><font size=2 color=  #fdfefe   face=verdana>" &  UCase(Nombre_Solapa)  & "</font></a><h2>" &  Nombre_Solapa   & "</h2><UL>"
						' FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"

						Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						ss7="<div title='" & UCase(Nombre_Solapa)& "'  id=s1solapa" & solord(x) & " style='display:none'><ul>"
						If rsx.eof  Or sFuncion="E" Then
							' sAccion=xSolasp & solord(x)
							sAccion =xSolaspd & solord(x) & "&NIV1=" & solord(x)  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5

							FTABELENUL=FTABELENUL & "<LI  ><font color=  #fdfefe   size=1 face=verdana ><a  target=WK1 href='" & sAccion &  "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"
						Else
							FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"

							Do While Not rsx.eof

								If eveant <> rsx(4) Then


									eveant  = rsx(4)
								End If


								sfidkD  = sfidkB
								stexto1 = rsx(2) & "." & rsx(6) & "."  & Capitalize(rsx(4) & "")  & "." & Capitalize(rsx(3) & "") & " - " & Capitalize(rsx(5) & "")
								stexto1 = Replace(stexto1," ","&nbsp;")

								Texto2 = rsx(2)
								wwsNIV2     = (rsx(0) & "")
								wwsNIV3     = "0"
								sAccion =xSolaspd & solord(x) & "&NIV1=" & solord(x)  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5

								'FTABELENUL=FTABELENUL & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>"

								ss7=ss7  & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>" '"<li>" &  stexto1 & "</li>"
								rsx.movenext




							Loop
						End If
						'	FTABELENUL=FTABELENUL & "</UL>"
						ss7=ss7 & "</UL></div>"
						'FTABELENUL=FTABELENUL  & ss7 & "</UL></LI>"
						Case 11 ' contenido

						Dim Inx
						Dim nrs(100)

						ReDim Preserve bsec(8)
						ReDim Preserve bnum(8)
						ReDim Preserve bnumG(8)


						Dim Pdef(13,18)
						' FTABELENUL= FTABELENUL & "<li class=solapa" & solord(x) & " ><a class=solapa" &solord(x) &" href='#'><i class='la la-cog'></i><font size=2 color=  #fdfefe   face=verdana>" &  UCase(Nombre_Solapa) & "</font></a><h2>" &  Nombre_Solapa  & solord(x)  & "pppppppppppppppppp</h2></li>"
						'  FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"
						strsql = "Select * From Objestr Where codobjeto =" & sCodigo
						Strsql = Strsql & " order by CONsecuen "
						'  response.write strsql & "<br>"
						Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						bsecuen = ""
						bDescri = " "
						bobj = ""

						iitemC=0
						ANTwwsNIV2=0
						strniv3=""
						ss11="<div title='" & UCase(Nombre_Solapa)& "' id=s1solapa" & solord(x) & " style='float: left'>"

						If  rsxx1.eof Then
							sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5
							'ss11=ss11  & "<LI class=cn1 data-id='solapa" & solord(x) & wwsNIV2 & "N3' ><font color=  #fdfefe   size=1 face=verdana ><a class=cn1 target=WK1 href='"& sAccion & "'>" & UCase(Nombre_Solapa) & "</a></font>"
							FTABELENUL=FTABELENUL & "<LI  class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  target=WK1 href='" & sAccion &  "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"
						Else
							FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"
							Do While Not rsxx1.eof
								ss11=ss11 & "<div id=s1solapa" & solord(x) & "N1" & "  style='float: left;display:none'><ul>"
								bsecuen = rsxx1(4)
								bDescri = rsxx1(3)
								bobj    = rsxx1(0)


								stexto1 = bsecuen & "." & Replace(Capitalize(bDescri & "")," ","&nbsp;")
								stexto1 = Replace(stexto1," ","&nbsp;")

								wwsNIV2     = bobj & ""
								wwsNIV3     = "0"

								sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5


								ss11=ss11  & "<LI class=cn1 data-id='solapa" & solord(x) & wwsNIV2 & "N3' ><font color=  #fdfefe   size=1 face=verdana ><a class=cn1 target=WK1 href='"& sAccion & "'>" & UCase(Replace(stexto1,"&nbsp;"," ")) & "</a></font>"
								wentro = 0

								strsql = "Select * From SysEstprty"
								Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								t = 0
								Do Until rs.eof
									For ti = 1 To 18
										If (rs(ti) & "") = "" Then
											sti = ""
										Else
											sti = rs(ti)
										End If

										Pdef(t,ti - 1) = sti

									Next
									rs.movenext
									t = t + 1
								Loop

								Call ofv.cerrarconsulta(rs)

								strsql = "Select * From ObjEstPrty where detcon = " & bobj
								'response.write strsql & "<br>"
								Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								t = 0
								Do Until rs.eof
									Select Case rs(4)
										Case "A"
										t = 0
										Case "B"
										t = 1
										Case "C"
										t = 2
										Case "D"
										t = 3
										Case "E"
										t = 4
										Case "F"
										t = 5
										Case "G"
										t = 6
									End Select

									For ti = 5 To 22
										If (rs(ti) & "") = "" Then
											sti = ""
										Else
											sti = rs(ti)
											Pdef(t,ti - 5) = sti
										End If



									Next
									rs.movenext
								Loop

								Call ofv.cerrarconsulta(rs)

								PInden = Pdef(0,7)
								If PInden <> "NO" Then PInden = "SI"

								PNumer = Pdef(0,12)
								If PNumer <> "NO" Then PNumer = "SI"

								PNumCls = Pdef(0,14)
								ReDim Preserve PNumTip(13)
								For i=0 To 13
									PNumTip(i) = Pdef(i,13)
									If PNumTip(i) = "" Then PNumTip(i) = "1"
								Next

								sant = "X"
								bseci = 0
								bsecix = 0

								For i=0 To 8
									bsec(i) = 0
									bnum(i) = 0
									bnumG(i) = 0
								Next

								inx = 0

								strsql = " SELECT * FROM CSESTDET "
								strsql = strsql & " where detcon = " & bobj & "  "
								strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
								strsql = strsql & " and not ( "
								strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
								strsql = strsql & " ) and stmodi = '' "

								If session("APLMDCLI")="BC" Then
									strsql=strsql & " AND DETTIPO < 'G'"
								End If
								strsql = strsql & " order by TSEC  "
								'	 response.write strsql

								Set Rsxx = Server.CreateObject("ADODB.Recordset")
								Rsxx.CursorType = 3
								Rsxx.CacheSize = 100

								Rsxx.Open strsql, cn


								If Not rsxx.eof Then
									ars = rsxx.GetRows()

								Else
									campos = rsxx.fields.count - 1
									ReDim  ars(CInt(campos),0)
									ars(39,0) = "N"
									wentro=1
								End If


								rsxx.Close
								Set rsxx = Nothing
								strniv3=""

								Do While wentro = 0


									strniv3 = strniv3 & "<div id='solapa" & solord(x) & bobj & "N3'" & " style='display:none'><ul>"

									For Inx = inx To UBound(ars,2)



										SCLAVE   = ars(0,inx)
										Stip     = ars(5,inx)

										'Secuencia
										Select Case stip
											Case "A"
											t = 7
											bseci = 0
											bsecix = 0
											bsec(1) = ars(4,inx)
											bnum(0) = bnum(0) + 1
											bnumG(0) = bnumG(0) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "B"
											t = 8
											bseci = 1
											bsecix = 1
											bsec(2) = ars(4,inx)
											bnum(1) = bnum(1) + 1
											bnumG(1) = bnumG(1) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "C"
											t = 9
											bseci = 2
											bsecix = 2
											bsec(3) = ars(4,inx)
											bnum(2) = bnum(2) + 1
											bnumG(2) = bnumG(2) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "D"
											t = 10
											bseci = 3
											bsecix = 3
											bsec(4) = ars(4,inx)
											bnum(3) = bnum(3) + 1
											bnumG(3) = bnumG(3) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "E"
											t = 11
											bseci = 4
											bsecix = 4
											bsec(5) = ars(4,inx)
											bnum(4) = bnum(4) + 1
											bnumG(4) = bnumG(4) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "F"
											t = 12
											bseci = 5
											bsecix = 5
											bsec(6) = ars(4,inx)
											bnum(5) = bnum(5) + 1
											bnumG(5) = bnumG(5) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "G"
											t = 13




											Select Case CStr(spsecn & "")
												Case "1"
												ta = 7
												bseci = 0
												Case "2"
												'ta = 8
												ta = 7
												bseci = 1
												Case "3"
												ta = 9
												bseci = 2
												Case "4"
												ta = 10
												bseci = 3
												Case "5"
												ta = 11
												bseci = 4
												Case "6"
												ta = 12
												bseci = 5
											End Select


											bnum(6) = bnum(6) + 1
											bnumG(6) = bnumG(6) + 1

											bsecix = bseci
											bsec(7) = ars(4,inx)

											Case "H"
											Select Case sant
												Case "A"
												bseci = 0
												Case "B"
												bseci = 1
												Case "C"
												bseci = 2
												Case "D"
												bseci = 3
												Case "E"
												bseci = 4
												Case "F"
												bseci = 5
											End Select
											bnum(7) = bnum(7) + 1
											bnumG(7) = bnumG(7) + 1
											bsecix = bseci + 1
											bsec(7) = ars(4,inx)
											Case "I"
											Select Case sant
												Case "A"
												bseci = 0
												Case "B"
												bseci = 1
												Case "C"
												bseci = 2
												Case "D"
												bseci = 3
												Case "E"
												bseci = 4
												Case "F"
												bseci = 5
											End Select
											bnum(8) = bnum(8) + 1
											bnumG(8) = bnumG(8) + 1
											bsecix = bseci + 1
											bsec(7) = ars(4,inx)
											Case "J"
											Select Case sant
												Case "A"
												bseci = 0
												Case "B"
												bseci = 1
												Case "C"
												bseci = 2
												Case "D"
												bseci = 3
												Case "E"
												bseci = 4
												Case "F"
												bseci = 5
											End Select
											bnum(8) = bnum(8) + 1
											bnumG(8) = bnumG(8) + 1
											bsecix = bseci + 1
											bsec(8) = ars(4,inx)

										End Select



										If stip <> "H" And stip <> "I"  And stip <> "J" Then

											For ti = 0 To 17
												Pdef(t,ti) = Pdef(t-7,ti)
											Next
											Pdef(t,18) = "SI"


											If ars(31,inx) = "NO" Then
												For ti = 11 To 28
													If (ars(ti,inx)&"") = "" Then
														sti = ""
													Else
														sti             = ars(ti,inx)
														Pdef(t,ti - 11) = sti
													End If



												Next
												Pdef(t,18) = "NO"

											End If
										End If

										BsecD = ""
										For i=1 To bsecix
											BsecD = BsecD & bsec(i) & "."
										Next





										'response.write strsql
										If Pdef(t,18) = "NO" Then
											tx = t
										Else
											tx = t - 7
										End If


										PNumCls = Pdef(tx,14)
										If PNumCls = "" Then PNumCls = "Simple"



										If stip <> "H" And stip <> "I"  And stip <> "J" Then
											PNumTip(tx) = Pdef(tx,13)
											If  PNumTip(tx) = "" Then PNumTip(tx) = "1"
										End If

										If stip = "G" Then
											If Pdef(tx,9) = "NO" Then
												MPvin = "NO"
											Else
												MPvin = "SI"
											End If
										End If

										' Anula la numeracion para este tipo o elemento
										If  PNumTip(tx) = "NO" Then PNumCls = PNumTip(tx)

										PFont  = Pdef(tx,0)
										PFsize = Pdef(tx,1)
										PFcol  = Pdef(tx,2)
										Pbgcol = Pdef(tx,3)
										PBorde = Pdef(tx,4)
										PFtipo = Pdef(tx,5)

										If marca = "CA" Then
											PFont  = "VERDANA"
											If smodi Then
												PFcol  = "#000080"
											Else
												PFcol  = "#aaaaaa"
											End If
											PFsize = "8"
											Pbgcol = ""
											PBorde = ""
											PFtipo = ""
										End If


										PSang  = Pdef(tx,6)
										txi = tx
										If tx > 6 Then txi = tx - 7
										PEspa  = Pdef(txi,8)



										Palign = Pdef(tx,16)
										If Palign = "" Then Palign = "left"
										Ptalig = Pdef(tx,17)


										txx = ta

										xmpvin = "NO"

										If stip = "G" And Pdef(txx,9) = "SI" Then
											xmpvin = "SI"
										End If

										If stip = "G" And MPvin  = "NO" Then
											xmpvin = "NO"
										End If

										xPNumS = "NO"

										If stip = "G" Then
											If Pnumer = "SI" Then
												Select Case UCase(Pnumcls)
													Case "SIMPLE"
													xPnumS = "SI"
													Case "COMPUESTA"
													xPnumS = "SI"
												End Select

											End If
										End If


										' TITULOS Y SUBTITULOS --------------------------------------------------

										If stip <> "G" And stip <> "H" And stip <> "I"  And stip <> "J" Then
											'Descripcion
											PNumS = ""
											If Pnumer = "SI" Then
												Select Case UCase(Pnumcls)
													Case "SIMPLE"
													Select Case PnumTip(tx)
														Case "1"
														BnumS = BnumG(bseci)
														Case Else
														BnumS = ConvNumer(BnumG(bseci),PnumTip(tx))
													End Select

													PnumS = BnumS & ". "
													Case "COMPUESTA"
													For i=0 To bseci
														Select Case PnumTip(i + 7)
															Case "NO"
															BnumS = ""
															Case "1"
															BnumS = BnumG(i) & "."
															Case Else
															BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
														End Select

														PnumS = PnumS & BnumS
													Next
													PnumS = PnumS & " "
												End Select

											End If



											'''''''''' arma tabla de indices '''''''''''''''''''''''''



											stexto1 = PnumS & Capitalize(ars(7,inx) & "")
											stexto1 = Replace(stexto1," ","&nbsp;")

											For nn=0 To (bseci+1)
												'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
												stexto1 = "&nbsp;" & stexto1
											Next

											wwsNIV2     = bobj & ""
											wwsNIV3     = (ars(0,inx) & "")

											sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

											'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
											iitemC=iitemC+1

											'Se saca border solid
											sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
											iitemC=iitemC+1

											sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
											Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

											'CTABELEN = ofv.GenCelda("","UT  style='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
											CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
											'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")
											wwwpad = "20"
											'if stip <> "A" then wwwpad = "30"
											wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:20px;background-color:" & sikon & ";padding-left:" & wwwpad & "px;'>" &  sLink & "<br></div>" & "<br>"
											strniv3=strniv3  & "<LI><font color=#fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>"

											'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




										End If



										' PARRAFOS DE TEXTO    --------------------------------------------------

										If stip = "G" Then

											PNumS = ""
											If Pnumer = "SI" Then
												Select Case UCase(Pnumcls)
													Case "NO"
													BnumG(6) = BnumG(6) - 1

													Case "SIMPLE"
													Select Case PnumTip(tx)
														Case "1"
														BnumS = BnumG(6)
														Case Else
														BnumS = ConvNumer(BnumG(6),PnumTip(tx))
													End Select

													PnumS = BnumS & ". "
													Case "COMPUESTA"
													For i=0 To 6
														If BnumG(i) > 0 Then
															Select Case PnumTip(i + 7)
																Case "NO"
																BnumS = ""
																Case "1"
																BnumS = BnumG(i) & "."
																Case Else
																BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
															End Select
															PnumS = PnumS & BnumS
														End If
													Next
													PnumS = PnumS & " "
												End Select

												If UCase(Pnumcls) <> "NO" Then

													sTextoV =  ars(8,inx)

													If Len(sTextoV) > 50 Then sTextoV =  Left(sTextoV,50) & "..."

													'''''''''' arma tabla de indices '''''''''''''''''''''''''



													stexto1 = PnumS & Capitalize(sTextoV & "")
													stexto1 = Replace(stexto1," ","&nbsp;")

													For nn=0 To (bseci+1)
														'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
														stexto1 = "&nbsp;" & stexto1
													Next

													wwsNIV2     = bobj & ""
													wwsNIV3     = (ars(0,inx) & "")

													sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

													'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
													iitemC=iitemC+1

													'Se saca border solid
													sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  &" style='border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
													iitemC=iitemC+1

													sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
													Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

													CTABELEN = ofv.GenCelda("","UT  style='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")

													'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
													wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "ttttt<br><br></div>" & "<br>"
													strniv3=strniv3  & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>"

													'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

												End If

											End If

										End If

										'------------------------------------------------


										Select Case stip
											Case "A"
											sant     = "A"
											sVinoAnt = sant

											Case "B"
											sant     = "B"
											sVinoAnt = sant
											Case "C"
											sant     = "C"
											sVinoAnt = sant
											Case "D"
											sant     = "D"
											sVinoAnt = sant
											Case "E"
											sant     = "E"
											sVinoAnt = sant
											Case "F"
											sant     = "F"
											sVinoAnt = sant
											Case Else

											sVinoAnt = stip

										End Select

									Next

									If inx > UBound(ars,2) Then
										Exit Do

									End If

									ssale = "no"
									Do Until ssale = "si"

										If     sant = "A" Then

											sant = "X"
										ElseIf sant = "B" Then

											sant = "A"
										ElseIf sant = "C" Then

											sant = "B"
										ElseIf sant = "D" Then

											sant = "C"
										ElseIf sant = "E" Then

											sant = "D"
										ElseIf sant = "F" Then

											sant = "E"
										End If

										If sant = "X" Then
											ssale = "si"

										End If

									Loop

								Loop
								strniv3= strniv3  & "</ul></div>"
								' strniv3="<ul><li class=solapa ><a class=solapa  href='#'><i class='la la-cog'></i><font size=2 color=  #fdfefe   face=verdana>" &  UCase(Nombre_Solapa) & "</font></a><h2>" &  Nombre_Solapa   & "</h2>"
								' strniv3=strniv3 & "<li>tyerter</li><li>tyerter</li><li>tyerter</li></li>tyerter</li></ul>"
								SS11=SS11 & strniv3 & "</LI>"
								'strniv3=""
								rsxx1.movenext
								SS11= SS11 & "</UL></DIV>"
							Loop

							' FTABELENUL=FTABELENUL  & "</LI>"
							SS11= SS11 & "</DIV>"
						End If
						Call ofv.cerrarconsulta(rsxx1)

						Case 9 ' tablas

						'Strsql = "SELECT *  "
						Strsql = "SELECT * "
						Strsql = Strsql & " from ObjTabDef Where codobjeto = " & sCodigo
						Strsql = Strsql & " order by tabsecuen "
						'  RESPONSE.WRITE STRSQL
						sInput=""

						Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						ss9="<div title='" & UCase(Nombre_Solapa)& "'  id=s1solapa" & solord(x) & " style='display:none'><ul>"
						If rsx.eof Then
							' sAccion=xSolasp & solord(x)
							sAccion =xSolaspd & solord(x) & "&NIV1=" & solord(x)  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5

							FTABELENUL=FTABELENUL & "<LI  ><font color=  #fdfefe   size=1 face=verdana ><a  target=WK1 href='" & sAccion &  "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"
						Else
							FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(Nombre_Solapa)  & "</a></font></LI>"

							Do While Not rsx.eof


								'sfidkD  = sfidkB
								stexto1 = Capitalize(rsx(4) & "")
								stexto1 = Replace(stexto1," ","&nbsp;")

								Texto2 = rsx(2)
								wwsNIV2     = (rsx(0) & "")
								wwsNIV3     = "0"
								sAccion =xSolaspd & solord(x) & "&NIV1=" & solord(x)  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5

								ss9=ss9  & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>" '"<li>" &  stexto1 & "</li>"
								rsx.movenext




							Loop
						End If
						'	FTABELENUL=FTABELENUL & "</UL>"
						ss9=ss9 & "</UL></div>"
						Call ofv.cerrarconsulta(rsx)
						Case Else
						sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5

						FTABELENUL=FTABELENUL & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 HREF='"& SACCION & "'> "  &  UCase(Nombre_Solapa) & "</a></font></LI>"
					End Select
					'FTABELENUL=FTABELENUL  &"</LI>"
					sinput2 = ""

					If sFuncion = "E" Or sFuncion = "ED"  Then

					Else
						'             xxxx = GetIndice(solord(x))
						'     sinput2 = ""
						sinput2 = sinput2 & "<div id='SOLX" & solord(x) & "' style='xdisplay:none; border-style:solid;border-width:1;border-color:#ff3333;position: absolute;top:1;left:1;' xxstyle='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 40px; border-style:solid;border-width:1px;border-color:grey;' >"
						sinput2 = sinput2' & GetIndice(solord(x))
						sinput2 = sinput2 & "</div>"

					End If


					CTABELEN = ofv.GenCelda("",  "uT align='center'" & wwwestilo  & " ","","","","","", sInput2  ,"si")
					FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
					'FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"
					'GetIndice(solord(x))
					'"uT align='center'"
					'& wwwestilo

					wwwk1 = wwwk1 & "<div  style='xdisplay:none;width=200px;height=200px;display: inline-block;margin: 30px 30px 30px 30px;' >" & sLink & "</div>"



				Else

					t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
					t3 = t3 & sInput
					t3 = t3 & "</A>"

					t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

					r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

				End If

				'response.write clng(sSClick) & " solact "  & clng(solord(x)) & chr(13)
				If CLng(sSClick) = CLng(solord(x)) Then
					session("codSol")=sSClick
					' response.write clng(sSClick)
					xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
					xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
					xverasp = xverasp & "&GLBSOLN=S" & solord(x)
					xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

					sHTML1 = ""

					sHTML1 = sHTML1 & "<form name=SSOLACT id=SSOLACT  action='" & xverasp & "' "
					sHTML1 = sHTML1 & " method=post target='"
					If CLng(sSClick) = 12 Then
						sHTML1 = sHTML1 & "_blank"
					Else
						sHTML1 = sHTML1 & "Principal"
					End If
					sHTML1 = sHTML1 & "' style='display:none;' >" & Chr(13)
					sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & Chr(13)
					sHTML1 = sHTML1 & "</form>" & Chr(13)

					'''''''''''''''''''''''''''''''''
					' NUEVO PARA VERSION 4.6

					'response.write request.querystring("Test") & " pregunta por Test " & "<br>"

					If request.querystring("Test") = "T" Then

						sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
						sHTML1 = sHTML1 & " window.top.CorSol='" & Replace(UCase(sSolapa),"*"," ") & "'; " & Chr(13)
						sHTML1 = sHTML1 & "</script>" & Chr(13)

					Else
						sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
						sHTML1 = sHTML1 & " window.top.CorSol='" & Replace(UCase(sSolapa),"*"," ") & "'; " & Chr(13)

						sHTML1 = sHTML1 & " var element =  document.getElementById(window.top.imenu);"

						sHTML1 = sHTML1 & " if (typeof(element) != 'undefined' && element != null)"
						sHTML1 = sHTML1 & " {"

						''''''''Agregado 10032015 - Version 4.6
						sHTML1 = sHTML1 & " element.style.backgroundColor='"&session("CMICTT")&"';"
						sHTML1 = sHTML1 & " element.style.color='"&session("CMICTTFT")&"';"

						''''''''Agregado 10032015 - FIN - Version 4.6

						sHTML1 = sHTML1 & " }"

						sHTML1 = sHTML1 & "</script>" & Chr(13)

					End If
					' NUEVO PARA VERSION 4.6 -FIN-
					'''''''''''''''''''''''''''''''''

					If session("WEBDESK") = "S" Then




					Else

						If  Session("FULLGUI") = "N" Then
							styl = "background-color:#888888;"
						Else
							styl = XspEx("#888888","#cccccc",0) & ";"
						End If

						xlfun = " onclick=""document.getElementById('MDBEDL').style.display=''; document.getElementById('DBEDTD1').style.display='none'; document.getElementById('DBEDTD2').style.display=''; document.getElementById('MM2').style.display=''; "" "
						xlfun2 = " onclick=""document.getElementById('MDBEDL').style.display='none'; document.getElementById('DBEDTD1').style.display=''; document.getElementById('DBEDTD2').style.display='none'; document.getElementById('MM2').style.display='none'; "" "

						macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:10px;' " & xlfun,"","90%","","","",Replace(UCase(sSolapa),"*"," "),"si")

						xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


						macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
						macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

						marow = ofv.GenRow("","","","",lengtd,"","",macel,"si")

						sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

					End If

				End If


			End If
		Next

		wwwk1 = wwwk1 & "</div>"

		CTABELEN = ofv.GenCelda("","uT align='center' ","","","","","",wwwk1,"si")
		FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
		'FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"

		If SMAINAPL = "OR" Or SMAINAPL = "PL" Then

		Else

			''-----------------------------------------------------------------------------------------------------

			''''''''''''''''''''''''''''''''''''''''''''
			'' Accesos directos del documento ''''''''''
			''''''''''''''''''''''''''''''''''''''''''''

			Dim ATABSCPT(200,2)
			Dim bConsultar()
			ISCPT = 0
			TSCPT = 0

			''    response.write "entro1"

			If sFuncion <> "X"   Then

				''  response.write "entro2"

				ReDim Preserve bConsultar(4)

				If sFuncion = "E"   Then

					'Se comenta para que no se muestren las normas relacionadas

					'bConsultar(2) = 0
					'bConsultar(4) = 0

					'bConsultar(1) = 1
					'bConsultar(3) = 1

					'else

					'bConsultar(2) = 0
					'bConsultar(4) = 0

					'bConsultar(1) = 0
					'bConsultar(3) = 1

				End If


				sVerEstado = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

				sverestado = sverestado & " and Objprty.Estado In (9,"
				For i = 1 To 4
					If bConsultar(i) Then sVerEstado = sVerEstado & i & ","
				Next
				sVerEstado = Left(sVerEstado, Len(sVerEstado) - 1) & ") "

				'''''''''''''''''''''''''''''''''''''''''



				sActcon  = "Consultar.asp?Codigo=NNNN&Tipo=TTTT"
				sActcon  = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
				sHref    = "&nbsp;"
				' response.write sActcon
				sFuncion = " onclick=" & Chr(34) & " window.top.abreWin('" & sActcon & "','TTTT'); return false;" & Chr(34)


				Strsql = "SELECT * FROM objacd  "
				Strsql = Strsql & " Where codobjeto=" & sCodigo
				Strsql = Strsql & " And tipobjeto = '" & stipo & "' "

				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				Do Until rs.eof

					''   response.write "entro3"


					If  rs(3) = "S" Then
						'''' Carpeta superior

						If  rs(6) = "S" Then


							''----------------

							strsql  = "select * from csbarr where codojasoc = " & sCodigo
							Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
							Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If  Not rs.eof Then
								essCodigo = rs1(1)
								essTipo   = rs1(2)

							Else
								essCodigo = 0
								essTipo   = "NO"
							End If
							Call ofv.cerrarconsulta(rs1)


							sNIVE   = "no"

							StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
							StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
							Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If Not rsIcono.EOF Then

								sIcono  = rsIcono(0)
								sTipoE  = rsIcono(1)
								sPerfil = rsIcono(2)
								sNIVE   = rsIcono(3)

							End If
							Call OFV.cerrarconsulta(rsIcono)

							If sNIVE   = "C"  Then

								strsql  = "select * from csbarr where codojasoc = " & essCodigo
								Strsql  = Strsql & " and tipobjasoc = '" & essTipo & "'"
								Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								If  Not rs.eof Then

									essCodigo = rs1(1)
									essTipo   = rs1(2)

								End If
								Call ofv.cerrarconsulta(rs1)


								'''''''''''''''''''''''''''''''''''''''''

								StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
								StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
								StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
								StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
								StrSql = StrSql & " where "
								StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
								StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
								StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
								StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
								StrSql = StrSql & "  objetos.objun =  1 )  and "
								StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
								StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
								StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
								''  response.write "<br>" & strsql & "<br>"
								sVino = "N"
								Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
								If  Not rs2.EOF Then

									Do Until rs2.EOF

										If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

											sVinook = "N"
											If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
												'' response.write "<br><br>entro"
												sVino   = "S"
												sVinook = "S"
											End If

											If SESSION("EMPMUN") = "S" Then
												StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
												StrSql = StrSql & " where "
												StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
												StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
												Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If  Not rsmas2.EOF Then
													sVino   = "S"
													sVinook = "S"
												End If
											End If

											sNIVE   = "D"

											If  sVinook = "S" Then

												stitem = rs2(4)
												StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
												StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
												Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If Not rsIcono.EOF Then
													sIcono  = rsIcono(0)
													sTipoE  = rsIcono(1)
													sPerfil = rsIcono(2)
													sNIVE   = rsIcono(3)

													sitem = sTipoE & ": " & rs2(1)

												End If
												Call OFV.cerrarconsulta(rsIcono)

												If sNIVE   = "D"  Then


													sSolapa =  sitem

													If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

														sInput = Replace(UCase(sSolapa),"*"," ")

													Else

														sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"

													End If

													sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
													sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))





													If session("WEBDESK") = "S" Then


														'sInput  = ">" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
														'sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
														'Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

														'wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

														'CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
														'FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")




													Else

														t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
														t3 = t3 & sInput
														t3 = t3 & "</A>"

														t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

														r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

													End If





												Else





												End If


											End If

										End If

										rs2.movenext
									Loop


								End If



								'''''''''''''''''''''''''''''''''''''''

							End If

							''----------------



						End If





					ElseIf  rs(3) = "P" Then
						'''' Carpeta primaria

						If  rs(6) = "S" Then

							strsql  = "select * from csbarr where codojasoc = " & sCodigo
							Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
							Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If  Not rs.eof Then
								essCodigo = rs1(1)
								essTipo   = rs1(2)

							Else
								essCodigo = 0
								essTipo   = "NO"
							End If
							Call ofv.cerrarconsulta(rs1)


							sNIVE   = "no"

							StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
							StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
							Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If Not rsIcono.EOF Then

								sIcono  = rsIcono(0)
								sTipoE  = rsIcono(1)
								sPerfil = rsIcono(2)
								sNIVE   = rsIcono(3)

							End If
							Call OFV.cerrarconsulta(rsIcono)




							StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
							StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
							StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
							StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
							StrSql = StrSql & " where "
							StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
							StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
							StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
							StrSql = StrSql & " (objetos.objun =  0" & session("UN") & " or "
							StrSql = StrSql & "  objetos.objun =  1 )  and "
							StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
							StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
							StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
							''   response.write "<br>" & strsql & "<br>"
							sVino = "N"
							Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If  Not rs2.EOF Then

								Do Until rs2.EOF

									If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

										sVinook = "N"
										If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
											'' response.write "<br><br>entro"
											sVino   = "S"
											sVinook = "S"
										End If

										If SESSION("EMPMUN") = "S" Then
											StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
											StrSql = StrSql & " where "
											StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
											StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
											Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If  Not rsmas2.EOF Then
												sVino   = "S"
												sVinook = "S"
											End If
										End If

										sNIVE   = "D"

										If  sVinook = "S" Then

											stitem = rs2(4)
											StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
											StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
											Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If Not rsIcono.EOF Then
												sIcono  = rsIcono(0)
												sTipoE  = rsIcono(1)
												sPerfil = rsIcono(2)
												sNIVE   = rsIcono(3)

												sitem = sTipoE & ": " & rs2(1)

											End If
											Call OFV.cerrarconsulta(rsIcono)

											If sNIVE   = "D"  Then


												sSolapa =  sitem

												If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
													sInput = Replace(UCase(sSolapa),"*"," ")
												Else
													sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
												End If

												sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
												sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))




												If session("WEBDESK") = "S" Then




													'    sInput  = "" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
													'    sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:normal;'  " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
													'    Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

													'    CTABELEN = ofv.GenCelda("","UT  style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' ","","","","","",sLink,"si")
													'    FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



													'' sInput  = ">" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
													sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
													Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

													wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

													CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
													FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
													'	FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"




												Else

													t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
													t3 = t3 & sInput
													t3 = t3 & "</A>"

													t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

													r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

												End If


											Else


											End If


										End If

									End If

									rs2.movenext
								Loop


								'''''''''''''''''''''''''''''''''''''''

							End If




						End If




					ElseIf  rs(3) = "R" Then
						'''' Carpeta dependiente

						If  rs(6) = "S" Then

							''----------------

							TSCPT = TSCPT + 1

							ATABSCPT(TSCPT,0) = ""
							ATABSCPT(TSCPT,1) = sCodigo
							ATABSCPT(TSCPT,2) = sTipo



							strsql  = "select * from csbarr where codojasoc = " & sCodigo
							Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
							Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If  Not rs.eof Then
								essCodigo = rs1(1)
								essTipo   = rs1(2)

							Else
								essCodigo = 0
								essTipo   = "NO"
							End If
							Call ofv.cerrarconsulta(rs1)


							sNIVE   = "no"

							StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
							StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
							Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If Not rsIcono.EOF Then

								sIcono  = rsIcono(0)
								sTipoE  = rsIcono(1)
								sPerfil = rsIcono(2)
								sNIVE   = rsIcono(3)

							End If
							Call OFV.cerrarconsulta(rsIcono)

							If sNIVE   = "C"  Then

								'  essCodigo = sCodigo
								'  essTipo   = sTipo

								'''''''''''''''''''''''''''''''''''''''''

								StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
								StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
								StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
								StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
								StrSql = StrSql & " where "
								StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
								StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
								StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
								StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
								StrSql = StrSql & "  objetos.objun =  1 )  and "
								StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
								StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
								StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
								''   response.write "<br>" & strsql & "<br>"
								sVino = "N"
								Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
								If  Not rs2.EOF Then

									Do Until rs2.EOF

										If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

											sVinook = "N"
											If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
												'' response.write "<br><br>entro"
												sVino   = "S"
												sVinook = "S"
											End If

											If SESSION("EMPMUN") = "S" Then
												StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
												StrSql = StrSql & " where "
												StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
												StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
												Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If  Not rsmas2.EOF Then
													sVino   = "S"
													sVinook = "S"
												End If
											End If

											sNIVE   = "D"

											If  sVinook = "S" Then

												stitem = rs2(4)
												StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
												StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
												Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If Not rsIcono.EOF Then
													sIcono  = rsIcono(0)
													sTipoE  = rsIcono(1)
													sPerfil = rsIcono(2)
													sNIVE   = rsIcono(3)

													sitem = sTipoE & ": " & rs2(1)

												End If
												Call OFV.cerrarconsulta(rsIcono)

												If sNIVE   = "D"  Then




												Else


													TSCPT = TSCPT + 1

													ATABSCPT(TSCPT,0) = ""
													ATABSCPT(TSCPT,1) = rs2(0)
													ATABSCPT(TSCPT,2) = rs2(2)


												End If


											End If

										End If

										rs2.movenext
									Loop

								End If



								Do While Not ISCPT > TSCPT

									If ISCPT = 0 Then ISCPT = 1


									For ISCPT=ISCPT To TSCPT



										StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
										StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
										StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
										StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
										StrSql = StrSql & " where "
										StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
										StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
										StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
										StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
										StrSql = StrSql & "  objetos.objun =  1 )  and "
										StrSql = StrSql & " RELAOBJ.codobjmaster = 0" & ATABSCPT(ISCPT,1)
										StrSql = StrSql & " and relaobj.tipobjmaster = '" & ATABSCPT(ISCPT,2) & "' " & sVerEstado
										StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
										''    response.write "<br>" & strsql & "<br>"
										sVino = "N"
										Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
										If  Not rs2.EOF Then

											Do Until rs2.EOF

												If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then



													sVinook = "N"
													If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
														'' response.write "<br><br>entro"
														sVino   = "S"
														sVinook = "S"
													End If

													If SESSION("EMPMUN") = "S" Then
														StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
														StrSql = StrSql & " where "
														StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
														StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
														Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
														If  Not rsmas2.EOF Then
															sVino   = "S"
															sVinook = "S"
														End If
													End If

													sNIVE   = "D"

													If  sVinook = "S" Then

														stitem = rs2(4)
														StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
														StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
														Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
														If Not rsIcono.EOF Then
															sIcono  = rsIcono(0)
															sTipoE  = rsIcono(1)
															sPerfil = rsIcono(2)
															sNIVE   = rsIcono(3)

															sitem = sTipoE & ": " & rs2(1)

														End If
														Call OFV.cerrarconsulta(rsIcono)

														If sNIVE   = "D" Then


															sSolapa =  sitem

															If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
																sInput = Replace(UCase(sSolapa),"*"," ")
															Else
																sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
															End If

															sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
															sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))



															If session("WEBDESK") = "S" Then



																'  sInput  = "" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
																'  sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:normal;'  " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
																'  Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

																'  CTABELEN = ofv.GenCelda("","UT  style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' ","","","","","",sLink,"si")
																'  FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")


																sInput  = ">" & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;") & ""
																sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
																Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

																wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

																CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
																FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
																' FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"

															Else

																t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
																t3 = t3 & sInput
																t3 = t3 & "</A>"

																t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

																r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

															End If



														Else



															TSCPT = TSCPT + 1

															ATABSCPT(TSCPT,0) = ""
															ATABSCPT(TSCPT,1) = rs2(0)
															ATABSCPT(TSCPT,2) = rs2(2)


														End If


													End If


												End If

												rs2.movenext
											Loop

										End If


									Next

								Loop



							End If



						End If


					Else

						''''' documentos o subcarpetas independientes


					End If


					''''''''''''''''''''''''''''


					'''''''''''''''''''''''''''''



					rs.movenext
				Loop

			End If

			''''''''''''''''''''''''''''''''''''''''''''


			''-----------------------------------------------------------------------------------------------------


		End If



		If session("WEBDESK") = "S" Then


			'''    sHTML = sHTML & Celda
			'''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

			'sHTML = sHTML & "<div id='WKTBSP1'  >"
			FTABELENUL=FTABELENUL & "</UL></NAV>" & ss11  & ss7 & ss9 &"</DIV>"
			'	sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "aaa  ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN , "si")
			'	shtml = sHTML &  FTABELENUL  & "</div>"

		Else

			'	sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"


			If  Session("FULLGUI") = "N" Then
				styl = "background-color:#888888;"
			Else
				styl = XspEx("#888888","#cccccc",0) & ";"
			End If

			styl = ""

			If ofv.ModTitCls = "1" Then
				'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			Else
				'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			End If

			sHTML = sHTML & "</div>"

		End If








		sHTML = sHTML & sHTML1
		sHTML = sHTML & sinput2


		Call ofv.cerrarconn(cn)

		sHTML = sHTML & "</body></html>"


	End If
	'If session("APLMDCLI")="BGAL" Then
	'Response.Write Replace(sHTML,"&nbsp;"," ") 'sHTML
	sHTML=""
	Whtml1 = ""
	Whtml1 = Whtml1 & "<script language=javascript >"
	'				Whtml1 = Whtml1 & " ClosIndice();" & Chr(13)

	Whtml1 = Whtml1 & "</script>"

	armasolapa2=FTABELENUL & Whtml1

	'Else
	'	Response.Write SHTML
	'End If

End Function

Function armasolapa(sCodigo,sTipo,sFuncionX)

	Dim sHTML
	Dim Celda
	ok=ofv.CheckUsuario()
	If ok <> "Y" Then
		uv = ofv.uv
		If uv = "NO" Then ok = ofv.ValidarUsuario
		sHTML = ok
	Else
		'response.expires=0
		bModificar = False
		ofv.ObtenerAtributos Session("Form2"),""
		bModificar = ofv.AModificar
		sCodigo=session("ncliente")
		sTipo=session("tcliente")

		' ver ale 9/4/2021
		' se toma por request.querystring ahora
		'if sCodigo="" then
		sTipo     = request.querystring("Tipo")
		sCodigo   = request.querystring("Codigo")
		' end if

		If session("APLMDCLI")="" Then
			aplcheck()
		End If

		If session("APLMDCLI")="BGAL" Then
			lengtd=""
		Else
			lengtd="20"
		End If
		iitemC=0
		If session("Nomodificar") = "S" Then bModificar = False

		Select Case session("ModMain")

			Case "ED"
			sFuncion="E"
			Case "CL"
			sFuncion="C"
			Case "WK"
			sFuncion="CD"



		End Select

		If sFuncion <> "" Then sFuncion=sFuncionX

		sSClick   = request.querystring("SSClick")
		sMenu     = request.querystring("Menu")
		sSector   = request.querystring("Sector")
		sNIV1     = request.querystring("NIV1")
		sNIV2     = request.querystring("NIV2")
		sNIV3     = request.querystring("NIV3")
		sNIV4     = request.querystring("NIV4")
		sNIV5     = request.querystring("NIV5")
		sMAINAPL  = request.querystring("MAINAPL")

		''Nuevo para Test
		sTest     = request.querystring("Test")

		' response.write sNIV1 & " &Test "

		''Nuevo para Test

		'response.write "recibe " & session("ModMain") & "<br>"

		Set cn = ofv.conectar(ofv.strconn2)
		Strsql = "Select TIPOCPT From TIPOBJETO Where TIPOBJETO ='" & sTipo & "'"
		Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		If Not rsx1.eof Then
			TIPOCPT=rsx1(0)
		Else
			TIPOCPT="D"
		End If
		Call ofv.cerrarconsulta(rsx1)


		'sHTML = "<html>"
		'sHTML = sHTML & "<HEAD><META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " content='text/html; charset=" & XLCHARSET & "'></HEAD>"

		wxs = "<script language=JavaScript>" & Chr(13)
		wxs = wxs & " var Fbtini = 0;" & Chr(13)
		wxs = wxs & " function IniWin() { " & Chr(13)
		wxs = wxs & "  if (document.getElementById('SSOLACT')) { " & Chr(13)
		wxs = wxs & " document.getElementById('SSOLACT').submit(); " & Chr(13)
		wxs = wxs & " } else { " & Chr(13)
		wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & Chr(13)
		wxs = wxs & " setTimeout (""IniWin()"", 100); } " & Chr(13)
		wxs = wxs & "  } " & Chr(13)
		wxs = wxs & " } " & Chr(13)
		wxs = wxs & " </script>" & Chr(13)

		'sHTML = sHTML & wxs

		'sHTML = sHTML & "<body onload='IniWin();' >"



		If sFuncion = "E" Then
			smodosol = "SP"
		Else
			smodosol = "TB"
		End If


		Dim stabla(20)
		Dim solord(20)
		Dim Nombres(20)
		Dim ASPs(20)
		Dim IKNs(20)
		Dim Solapas(20)
		Dim Solapas2(20)
		Dim SolapasI(20)
		Dim rs

		ASPs(0) = "CoWKFselect.asp?SKIND=1"

		wwwsaltador = "<BR>"


		If SMAINAPL = "OR" Or SMAINAPL = "PL" Then



			ASPs(0) = "COWKFVERIFICACION.asp?SKIND=1"

			solord(0) = 0
			solord(1) = 1
			solord(2) = 2
			solord(3) = 3
			solord(4) = 4
			solord(5) = 5
			solord(6) = 6
			solord(7) = 7
			solord(8) = 8
			solord(9) = 9
			solord(10) = 10
			solord(11) = 11
			solord(12) = 12
			solord(13) = 13
			solord(14) = 14
			solord(15) = 15
			solord(16) = 16
			solord(17) = 17
			solord(18) = 18
			solord(19) = 19
			solord(20) = 20



			perfilsp = "S"

			If SMAINAPL = "OR" Then
				sEsquema = "CIRCUITO AUTOEVALUACION"
			Else
				sEsquema = "PLAN DE CONTINGENCIAS"
			End If

			xestado = 1

			Strsql = "Select estado,CODOBJETO,TIPOBJETO,RELLINK From ORRSKCIRCUITO Where OBJID = 0" & sCodigo
			Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsx1.eof Then
				xestado = rsx1(0)
				sCodigoVINC = rsx1(1)
				sTIPOVINC = rsx1(2)
				sRELVINC = rsx1(3)
			End If
			Call ofv.cerrarconsulta(rsx1)


			If (sFuncion = "CD") Then
				ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=0"
			Else
				ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=1"
			End If




			For z=0 To 20
				Solapas(z) = "N"
			Next


			For xx = 1 To 20
				Nombres(xx) = "N"
			Next

			'Solapas(1)  = REPLACE("DETALLE DEL CIRCUITO"," ","*")
			Solapas(1)  = Replace("MATRIZ DE AUTOEVALUACION"," ","*")
			Solapas2(1) = ASPs(1)
			xstrsol      = 1
			xfinal       = 1
			'''   sSClick      = 99



			If session("ModMain") = "WK"  Then
				Solapas(0)  = "WORKFLOW"
				Solapas2(0) = ASPs(0)
				SolapasI(0) = "#e7ea1c"
				xstrsol     = 0

			Else
				xstrsol     = 1
			End If



			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		Else


			ASPs(1) = "edPropiedades.asp?SKIND=1"
			IKNs(1) = "#baffc9"

			If sFuncion = "E" Then
				''' ASPs(2) = "edConsideraciones.asp?SKIND=1"
				ASPs(2) = "EDNWEDT.asp?SKIND=1"
			Else
				'''    ASPs(2) = "CoConsideraciones.asp?SKIND=1"
				ASPs(2) = "CONWEDT.asp?SKIND=1"
			End If

			IKNs(2) = "#ffffba"

			If sFuncion = "E" Then
				ASPs(3) = "edSectores.asp?SKIND=1"
			Else
				ASPs(3) = "CoSectores.asp?SKIND=1"
			End If

			IKNs(3) = "#bae1ff"



			ASPs(4) = "edEventos.asp?SKIND=1"

			IKNs(4) = "#ff7f50"

			If sFuncion = "E" Then
				ASPs(5) = "edProcedimientos.asp?SKIND=1"
			Else
				ASPs(5) = "CoProcedimientos.asp?SKIND=1"
			End If

			IKNs(5) = "#ffdfba"

			If sFuncion = "E" Then
				ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
			Else
				ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
			End If

			IKNs(16) = "#ffb3ba"

			If sFuncion = "E" Then
				ASPs(6) = "edObjDsc.asp?SKIND=1"
			Else
				ASPs(6) = "CoObjDsc.asp?SKIND=1"
			End If

			IKNs(6) = "#3dece2"

			If sFuncion = "E" Then
				ASPs(7) = "edControles.asp?SKIND=1"
			Else
				ASPs(7) = "CoControles.asp?SKIND=1"
			End If

			IKNs(7) = "#e397e2"

			If sFuncion = "E" Then
				ASPs(8) = "edEsquemas.asp?SKIND=1"
			Else
				ASPs(8) = "CoEsquemas.asp?SKIND=1"
			End If

			IKNs(8) = "#f6d2d2"

			If sFuncion = "E" Then
				ASPs(9) = "edTablas.asp?SKIND=1"
			Else
				ASPs(9) = "CoTablasCons.asp?SKIND=1"
			End If

			IKNs(9) = "#97e3c6"

			If sFuncion = "E" Then
				ASPs(10) = "edarchivos.asp?SKIND=1"
			Else
				ASPs(10) = "CoArchCons.asp?SKIND=1"
			End If

			IKNs(10) = "#ddc37a"

			If sFuncion = "E" Then
				ASPs(11) = "edestr.asp?SKIND=1"
			Else
				''  ASPs(11) = "Coestrex.asp?SKIND=1"
				ASPs(11) = "CoestrCons.asp?SKIND=1"
			End If

			IKNs(11) = "#c0db50"

			If sFuncion = "E" Then
				ASPs(12) = "CursogramaLnch.asp?SKIND=1"
				'''   ASPs(12) = "edProcedimientos.asp?SKIND=1"
			Else
				ASPs(12) = "CursogramaLnch.asp?SKIND=1"
				''''' ASPs(12)="CursogramaView.asp?SKIND=2"
			End If

			IKNs(12) = "#66cdaa"

			sMAINAPL = request.querystring("MAINAPL")
			If sFuncion = "E" Then
				ASPs(13) = "edarchivos.asp?SKIND=2"
			Else
				ASPs(13) = "CoArchCons.asp?SKIND=2"
			End If

			''''  IKNs(13) = "#ddffc9"


			If sFuncion = "E" Then
				ASPs(14) = "edFAQ.asp?SKIND=1"
			Else
				ASPs(14) = "CoFAQ.asp?SKIND=1"
			End If

			'---- IKNs(14) = "#ddffc9"


			If sFuncion = "E" Then


				ASPs(15) = "edGeneral.asp?SKIND=1"
			Else


				If TIPOCPT="K" Then
					ASPs(15) = "CoGeneralForm.asp?SKIND=1"
				Else
					ASPs(15) = "CoGeneral.asp?SKIND=1"
				End If
				' ASPs(15) = "CoGeneral.asp?SKIND=1"
			End If

			IKNs(15) = "#ddffc9"

			If sFuncion = "E" Then
				ASPs(20) = "edportada.asp?SKIND=1"
			Else
				ASPs(20) = "Coportada.asp?SKIND=1"
			End If

			IKNs(20) = "#f6d2d2"

			solord(0) = 0
			solord(1) = 15
			solord(2) = 1
			solord(3) = 6
			solord(4) = 3
			solord(5) = 2
			solord(6) = 4
			solord(7) = 5
			solord(8) = 7
			solord(9) = 8
			solord(10) = 11
			solord(11) = 10
			solord(12) = 9
			solord(13) = 12
			solord(14) = 13
			solord(15) = 14
			solord(16) = 16
			solord(17) = 17
			solord(18) = 18
			solord(19) = 19
			solord(20) = 20




			perfilsp = "S"


			siconmas = "ICorAcciones.jpg"
			Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
			'	response.write strsql & "<br>"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs.eof Then
				sEsquema = rs(1)
				sPerfil = rs(3)
				COMENTARIOS=RS("COMENTARIOS")
			End If

			Call ofv.cerrarconsulta(rs)


			For z=0 To 20
				Solapas(z) = "N"
			Next




			Strsql = "Select * From Objperfildoc Where codobjeto =0" & sCodigo & " and tipobjeto ='" & stipo & "' "
			'	  response.write strsql
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If rs.eof Then
				perfilsp = "N"
			End If

			If perfilsp = "N" Then
				Call ofv.cerrarconsulta(rs)
				Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"

				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			End If

			If Session("Segco") <> "S" Then
				For xx = 1 To 20
					Nombres(xx) = "N"
				Next

				Solapas(15)  = Replace(RS(31)," ","*")
				Solapas2(15) = ASPs(15)
				SolapasI(15) = IKNs(15)
				xstrsol      = 1
				xfinal       = 1
				sSClick      = 99


			Else

				For xx = 1 To 20
					solord(xx) = 17
					Nombres(xx) = "N"
				Next



				For xx = 2 To 16
					stabla (xx - 1) = rs(xx + 30)
				Next



				yy = 0
				For xx = 1 To 15

					If CInt(stabla (xx)) < 99 Then
						yy = yy + 1
						'   response.write stabla (xx) & "=" &  xx & chr(13)
						solord (stabla (xx)) = xx
					End If
				Next
				tyy = yy



				'''  if session("ModMain") = "WK" and sFuncion = "E" then
				If session("ModMain") = "WK"  Then
					'  if session("ModMain") = "WK" and ( session("MODSTART") <> "XP" or (session("MODSTART") = "XP" and session("WKFCODIGO") = sCodigo) ) then
					Solapas(0)  = "WORKFLOW"
					Solapas2(0) = ASPs(0)
					SolapasI(0) = "#e7ea1c"
					''' xstrsol     = 0
					xstrsol     = 0

					If session("APLMDCLI") <> "BC" Then


						' se saco porque generaba error con la solapa propiedades en el primer lugar 22_9_2017

						'       if sSClick      = 99 then sSClick      = solord(1)

					End If
				Else
					xstrsol     = 1
				End If





				X=0

				If  sFuncion <> "E"  Then

					'if sFuncion <> "E" then

					''' Solapas(20)  = "PORTADA"
					Solapas2(20) = ASPs(20)
					SolapasI(20) = IKNs(20)

					Strsql = "SELECT OBJportada.detcon, OBJportada.detsecuen, OBJportada.dettipo from OBJportada "
					Strsql = Strsql & " where OBJportada.codobjeto=" & sCodigo & " and OBJportada.detcon > 0 and OBJportada.dettipo <> 'N' "

					' response.write strsql
					Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If Not rsxx1.eof Then

					End If

					For xiord=0 To 19

						If solord(xiord) = 1 Then

							xmovey = 20

							'  for xiord2=xiord + 1 to 19
							'      xmoven           = solord(xiord2)
							'      solord(xiord2)   = xmovey
							'      xmovey           = xmoven
							'  next



							Exit For
						End If
					Next


					WWWMOVER = 20

					For xiord=1 To 20

						XXXMOVER      = solord(xiord)
						solord(xiord) = WWWMOVER
						WWWMOVER      = XXXMOVER

					Next




				End If

				For z = 2 To 16

					If rs(z) Then




						If (sFuncion = "E"  ) Or (sFuncion = "C" And z <> 2 And z <> 4 And z <> 5 ) Or (sFuncion = "CD" And z <> 2 And z <> 5 )  Or (sFuncion = "D"  And z = 4 ) Then
							x = x + 1
							Solapas(z-1)  = Replace(RS(z + 15)," ","*")
							Solapas2(z-1) = ASPs(z-1)
							SolapasI(z-1) = IKNs(z-1)

							If sFuncion <> "E" And z = 6 Then

								Solapas(16)  = "CIRCUITO COMPLETO"
								Solapas2(16) = ASPs(16)
								SolapasI(16) = IKNs(16)

								For xiord=0 To 19

									If solord(xiord) = 5 Then

										xmovey = 16

										For xiord2=xiord + 1 To 19
											xmoven           = solord(xiord2)
											solord(xiord2)   = xmovey
											xmovey           = xmoven
										Next



										Exit For
									End If
								Next


							End If

						End If
					End If

				Next


				xfinal = UBound(Solapas)

				' response.write "sSClick=" & sSClick & " solord(1) " &  solord(1)& chr(13)
				'  if sSClick      = 1 then sSClick      = solord(1)
			End If

			''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

			Call ofv.cerrarconsulta(rs)


		End If


		TABELEN = ""
		FTABELEN = ""
		Dim FTABELENUL
		Dim ars
		'FTABELENUL= "<div id=menucor data-cli='"&session("APLMDCLI")& "'><nav><h2><i class='la la-bars'></i>SOLAPAS</h2><UL>"
        FTABELENUL="<div id=contm STYLE='overflow:hidden;' xclass=sidebar><ul class=ulmenu STYLE= 'overflow:auto;Height:100%;' id=myMenu >" & chr(13)
        FTABELENUL=FTABELENUL & chr(13) & "<li><div class='item-container'><div class='item-text'><img src='" & Session("CliLogo") &"'><hr></div></div></li>" & Chr(13)

		xSolasp = ""
		xSolasp = xSolasp & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
		xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
		xSolasp = xSolasp & "&Funcion=" & sFuncion
		xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
		xSolasp = xSolasp & "&SSCLICK="


		xSolaspd = ""
		xSolaspd = xSolaspd & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
		xSolaspd = xSolaspd & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
		xSolaspd = xSolaspd & "&Funcion=" & sFuncion
		xSolaspd = xSolaspd & "&SSCLICK="

		r1 = ""


		If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
			xGenConta = ofv.GetTit(CLng(ofv.ModTitCls) + 5)
			xGenConta = ofv.FTitDat1

			xsclase = "xts"

		Else



			xsclase = "xts"

		End If

		xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
		xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"



		xsolini = 99


		If session("WEBDESK") = "S" And session("ModMain") = "WK" Then

			sInput  = "<font color='#990000' >"
			sInput  = sInput & "Este documento requiere su confirmacion en el workflow!!!</font>"
			sLink   = sInput

			CTABELEN = ofv.GenCelda("","vT  style='font-size:8pt;background-color:#ffffff;border-style:solid;border-width:1;border-color:#000000;padding-left:5px;' ","","","","","",sLink,"si")
			'  FTABELEN = FTABELEN & ofv.GenRow("","","","","25","","",CTABELEN,"si")

		End If

		If session("WEBDESK") = "S" And session("ModMain") <> "WK" And (sFuncion = "C" Or sFuncion = "CD") Then

			xverasp = "CoObjRecomendar.asp?Tipobjeto=" & sTipo & "&Codobjeto="
			xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
			xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5


			sInput  = "<b><font color='#ff0000' size=2  >&#058;&#0183;)</font></b>&nbsp;&nbsp;"
			'''  sInput  = ""
			sInput  = sInput & "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >Recomendar </font>"
			'''  sInput  = sInput & "&nbsp;<font color='#ff0000' size=2 >&#033;</font>"
			sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:bold;'  ", xverasp, "Principal", "", "", sInput, "si")
			Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

			wwwestilo = " style='background-color:#FFEEDD;border-style:dotted;border-width:2;border-color:#990000;padding-left:10px;' "

			CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
			' FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")

		End If

		wwwk1 = "<div id='portadaFR' name='contenedor' class='portada' style='display:none;' >"

		'response.write   xstrsol & " to  "  & xfinal &  chr(13)
		For x = xstrsol To xfinal
			'' for x = xstrsol to 2
			'response.write solapas(solord(x)) & " x=" & x &  chr(13)
			'response.write  solapas(solord(x)) & "<br>"
			If solapas(solord(x)) <> "N" Then
				sSolapa =  solapas(solord(x))
				sikon   =  SolapasI(solord(x))


				If  xsolini = 99 Then  xsolini = solord(x)

				If  sSClick = 99 Then  sSClick = xsolini

				If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


					If sFuncion = "E" Or sFuncion = "ED"  Then
						sInput  = Replace(UCase(sSolapa),"*"," ")
					Else
						sInput  = "<img src='" & sikon & "' >"

					End If




					xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:10px;'"


				Else


					If sFuncion = "E" Or sFuncion = "ED"  Then
						sInput  = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
					Else
						'''  sInput  = "<img width='80px' src='" & sikon & "' >"
						sInput  = "<input type='button' value="'" & replace(Ucase(sSolapa),"*"," ") & "'" style='font-weight:bold; font-size:16px; background-color: #baffc9; border-radius: 60px; width: 167px; height: 40px'>"
					End If




					xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:10px;'"

				End If

				If session("WEBDESK") = "S" Then


					Dim Nombre_Solapa
					If sFuncion = "E" Or sFuncion = "ED"  Then
						sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >"
						sInput  = sinput & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;")
						sInput  = sinput & "</font>"
						Nombre_Solapa = Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," "," ")
					Else
						sInput = ""


						Nombre_Solapa = Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," "," ")

						sInput  = "<div id='LNK" & solord(x) & "'  valign= 'middle' style='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 50px; border-style:solid; border-width:1px; border-color:grey; display:flex; justify-content:center; align-content:center; flex-direction:column;'>qqqqqqqqqqqqqqqqqqqqq" & Nombre_Solapa & "</div>"
					End If


					'   ssMnnSOL = "MNNSOL" & X
					iitemC=iitemC+1
					'   sLink   = "<span id='" & ssMnnSOL & "'  class='XX' >" & sInput & "</span> "
					'   sLink   = OFV.GenLink("", "cmb  onclick=" & chr(34) & "SetMnnSOL('"&ssMnnSOL&"');" & chr(34) & "style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")


					If sFuncion = "E" Or sFuncion = "ED"  Then
						sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput , "si")

					Else
						sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput , "si")
					End If

					'     sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput , "si")
					'
					Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink  & "&nbsp;&nbsp;</span> "

					wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

					If (solord(x)) = 0 Then
						wwwestilo = " style='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
					End If

					CTABELEN = ofv.GenCelda("","uT align='center' " & wwwestilo & " ","","","","","",sLink,"si")
					FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

					SLI=""
					Select Case solord(x)



						Case 5

						eveant = ""

						Strsql = "SELECT *  "
						Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & sCodigo
						strsql = strsql & " order by secuen, secuencia,codproced "
						'  RESPONSE.WRITE STRSQL
						sInput=""

						Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						ss7=""
						If rsx.eof  Then
							xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
							xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
							xverasp = xverasp & "&GLBSOLN=S" & solord(x)
							xverasp = xverasp & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5 '& "&ddd=" & sFuncion

							FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=Principal href='" & xverasp &  "'>"  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"
						Else
						    ss7="<ul>"
							xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
							xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
							xverasp = xverasp & "&GLBSOLN=S" & solord(x)
							xverasp = xverasp & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5 '& "&ddd=" & sFuncion
							If sFuncion = "E" Then
							FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a class=ntmnn target=Principal data-id='s1solapa" & solord(x) & "' class=""link"" href='"&xverasp & "'>"  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"
						   else
							FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a class=ntmnn target=Principal data-id='s1solapa" & solord(x) & "' class=""link"" xhref='"&xverasp & "'>"  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"
							end if

							Do While Not rsx.eof

								If eveant <> rsx(4) Then


									eveant  = rsx(4)
								End If


								sfidkD  = sfidkB
								stexto1 = rsx(2) & "." & rsx(6) & "."  & Capitalize(rsx(4) & "")  & "." & Capitalize(rsx(3) & "") & " - " & Capitalize(rsx(5) & "")
								stexto1 = Replace(stexto1," ","&nbsp;")

								Texto2 = rsx(2)
								wwsNIV2     = (rsx(0) & "")
								wwsNIV3     = "0"
								sAccion =xSolaspd & solord(x) & "&NIV1=" & solord(x)  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5

                                xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
								xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
								xverasp = xverasp & "&GLBSOLN=S" & solord(x)
								xverasp = xverasp & "&NIV1=" & solord(x)  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5 '& "&ddd=" & sFuncion


							'	ss7=ss7  & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>" '"<li>" &  stexto1 & "</li>"
							    ss7=ss7  & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=Principal href='" & xverasp &  "'> "  &  stexto1 & "</a></div></div></li>"

								rsx.movenext

							Loop
							  ss7=ss7 & "</UL>"
						End If


						If sFuncion = "E" Then
						 FTABELENUL=FTABELENUL  & "</LI>"
						else
						 FTABELENUL=FTABELENUL  & ss7 & "</LI>"
						end if
						Case 11 ' contenido

						Dim Inx
						Dim nrs(100)

						ReDim Preserve bsec(8)
						ReDim Preserve bnum(8)
						ReDim Preserve bnumG(8)


						Dim Pdef(13,18)
						strsql = "Select * From Objestr Where codobjeto =" & sCodigo
						Strsql = Strsql & " order by CONsecuen "
						'response.write strsql
						Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						bsecuen = ""
						bDescri = " "
						bobj = ""

						iitemC=0
						ANTwwsNIV2=0
						strniv3=""
					'	ss11="<div title='" & UCase(Nombre_Solapa)& "' id=s1solapa" & solord(x) & " style='float: left'>"
                        ss11="<ul>"
						If  rsxx1.eof Then
							sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5
							'ss11=ss11  & "<LI class=cn1 data-id='solapa" & solord(x) & wwsNIV2 & "N3' ><font color=  #fdfefe   size=1 face=verdana ><a class=cn1 target=WK1 href='"& sAccion & "'>" & UCase(Nombre_Solapa) & "</a></font>"
							FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text'  data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=WK1 href='" & sAccion &  "'>"  &  UCase(Nombre_Solapa)  & "</a></div></div></li>"
						Else
							xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
							xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
							xverasp = xverasp & "&GLBSOLN=S" & solord(x)
							xverasp = xverasp & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5 '& "&ddd=" & sFuncion
                            If sFuncion = "E" Then
                            FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a class=ntmnn target=Principal data-id='s1solapa" & solord(x) & "' class=""link"" href='"&xverasp & "'>"  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"
                           else
							FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a class=ntmnn target=Principal data-id='s1solapa" & solord(x) & "' class=""link"" xhref='"&xverasp & "'>"  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"
							end if
							Do While Not rsxx1.eof
								'ss11=ss11 & "<div id=s1solapa" & solord(x) & "N1" & "  style='float: left;display:none'><ul>"
								bsecuen = rsxx1(4)
								bDescri = rsxx1(3)
								bobj    = rsxx1(0)


								stexto1 = bsecuen & "." & Replace(Capitalize(bDescri & "")," ","&nbsp;")
								stexto1 = Replace(stexto1," ","&nbsp;")

								wwsNIV2     = bobj & ""
								wwsNIV3     = "0"

								'sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5
                                xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
								xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
								xverasp = xverasp & "&GLBSOLN=S" & solord(x)
								xverasp = xverasp & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5 '& "&ddd=" & sFuncion

                                ss11=ss11 & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=Principal href='" & xverasp &  "'> "  &  stexto1 & "</a></div></div></li>"
								'ss11=ss11  & "<LI class=cn1 data-id='solapa" & solord(x) & wwsNIV2 & "N3' ><font color=  #fdfefe   size=1 face=verdana ><a class=cn1 target=WK1 href='"& sAccion & "'>" & UCase(Replace(stexto1,"&nbsp;"," ")) & "</a></font>"
								wentro = 0

								strsql = "Select * From SysEstprty"
								Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								t = 0
								Do Until rs.eof
									For ti = 1 To 18
										If (rs(ti) & "") = "" Then
											sti = ""
										Else
											sti = rs(ti)
										End If

										Pdef(t,ti - 1) = sti

									Next
									rs.movenext
									t = t + 1
								Loop

								Call ofv.cerrarconsulta(rs)

								strsql = "Select * From ObjEstPrty where detcon = " & bobj
								'response.write strsql & "<br>"
								Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								t = 0
								Do Until rs.eof
									Select Case rs(4)
										Case "A"
										t = 0
										Case "B"
										t = 1
										Case "C"
										t = 2
										Case "D"
										t = 3
										Case "E"
										t = 4
										Case "F"
										t = 5
										Case "G"
										t = 6
									End Select

									For ti = 5 To 22
										If (rs(ti) & "") = "" Then
											sti = ""
										Else
											sti = rs(ti)
											Pdef(t,ti - 5) = sti
										End If



									Next
									rs.movenext
								Loop

								Call ofv.cerrarconsulta(rs)

								PInden = Pdef(0,7)
								If PInden <> "NO" Then PInden = "SI"

								PNumer = Pdef(0,12)
								If PNumer <> "NO" Then PNumer = "SI"

								PNumCls = Pdef(0,14)
								ReDim Preserve PNumTip(13)
								For i=0 To 13
									PNumTip(i) = Pdef(i,13)
									If PNumTip(i) = "" Then PNumTip(i) = "1"
								Next

								sant = "X"
								bseci = 0
								bsecix = 0

								For i=0 To 8
									bsec(i) = 0
									bnum(i) = 0
									bnumG(i) = 0
								Next

								inx = 0

								strsql = " SELECT * FROM CSESTDET "
								strsql = strsql & " where detcon = " & bobj & "  "
								strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
								strsql = strsql & " and not ( "
								strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
								strsql = strsql & " ) and stmodi = '' "

								If session("APLMDCLI")="BC" Then
									strsql=strsql & " AND DETTIPO < 'G'"
								End If
								strsql = strsql & " order by TSEC  "
							'		 response.write strsql

								Set Rsxx = Server.CreateObject("ADODB.Recordset")
								Rsxx.CursorType = 3
								Rsxx.CacheSize = 100

								Rsxx.Open strsql, cn


								If Not rsxx.eof Then
									ars = rsxx.GetRows()

								Else
									campos = rsxx.fields.count - 1
									ReDim  ars(CInt(campos),0)
									ars(39,0) = "N"
									wentro=1
								End If


								rsxx.Close
								Set rsxx = Nothing
								strniv3=""

								Do While wentro = 0


									strniv3 = strniv3 & "<ul>"

									For Inx = inx To UBound(ars,2)



										SCLAVE   = ars(0,inx)
										Stip     = ars(5,inx)

										'Secuencia
										Select Case stip
											Case "A"
											t = 7
											bseci = 0
											bsecix = 0
											bsec(1) = ars(4,inx)
											bnum(0) = bnum(0) + 1
											bnumG(0) = bnumG(0) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "B"
											t = 8
											bseci = 1
											bsecix = 1
											bsec(2) = ars(4,inx)
											bnum(1) = bnum(1) + 1
											bnumG(1) = bnumG(1) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "C"
											t = 9
											bseci = 2
											bsecix = 2
											bsec(3) = ars(4,inx)
											bnum(2) = bnum(2) + 1
											bnumG(2) = bnumG(2) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "D"
											t = 10
											bseci = 3
											bsecix = 3
											bsec(4) = ars(4,inx)
											bnum(3) = bnum(3) + 1
											bnumG(3) = bnumG(3) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "E"
											t = 11
											bseci = 4
											bsecix = 4
											bsec(5) = ars(4,inx)
											bnum(4) = bnum(4) + 1
											bnumG(4) = bnumG(4) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "F"
											t = 12
											bseci = 5
											bsecix = 5
											bsec(6) = ars(4,inx)
											bnum(5) = bnum(5) + 1
											bnumG(5) = bnumG(5) + 1
											For i=bseci + 1 To 8
												bnum(i) = 0
												bnumG(i) = 0
											Next
											Case "G"
											t = 13




											Select Case CStr(spsecn & "")
												Case "1"
												ta = 7
												bseci = 0
												Case "2"
												'ta = 8
												ta = 7
												bseci = 1
												Case "3"
												ta = 9
												bseci = 2
												Case "4"
												ta = 10
												bseci = 3
												Case "5"
												ta = 11
												bseci = 4
												Case "6"
												ta = 12
												bseci = 5
											End Select


											bnum(6) = bnum(6) + 1
											bnumG(6) = bnumG(6) + 1

											bsecix = bseci
											bsec(7) = ars(4,inx)

											Case "H"
											Select Case sant
												Case "A"
												bseci = 0
												Case "B"
												bseci = 1
												Case "C"
												bseci = 2
												Case "D"
												bseci = 3
												Case "E"
												bseci = 4
												Case "F"
												bseci = 5
											End Select
											bnum(7) = bnum(7) + 1
											bnumG(7) = bnumG(7) + 1
											bsecix = bseci + 1
											bsec(7) = ars(4,inx)
											Case "I"
											Select Case sant
												Case "A"
												bseci = 0
												Case "B"
												bseci = 1
												Case "C"
												bseci = 2
												Case "D"
												bseci = 3
												Case "E"
												bseci = 4
												Case "F"
												bseci = 5
											End Select
											bnum(8) = bnum(8) + 1
											bnumG(8) = bnumG(8) + 1
											bsecix = bseci + 1
											bsec(7) = ars(4,inx)
											Case "J"
											Select Case sant
												Case "A"
												bseci = 0
												Case "B"
												bseci = 1
												Case "C"
												bseci = 2
												Case "D"
												bseci = 3
												Case "E"
												bseci = 4
												Case "F"
												bseci = 5
											End Select
											bnum(8) = bnum(8) + 1
											bnumG(8) = bnumG(8) + 1
											bsecix = bseci + 1
											bsec(8) = ars(4,inx)

										End Select



										If stip <> "H" And stip <> "I"  And stip <> "J" Then

											For ti = 0 To 17
												Pdef(t,ti) = Pdef(t-7,ti)
											Next
											Pdef(t,18) = "SI"


											If ars(31,inx) = "NO" Then
												For ti = 11 To 28
													If (ars(ti,inx)&"") = "" Then
														sti = ""
													Else
														sti             = ars(ti,inx)
														Pdef(t,ti - 11) = sti
													End If



												Next
												Pdef(t,18) = "NO"

											End If
										End If

										BsecD = ""
										For i=1 To bsecix
											BsecD = BsecD & bsec(i) & "."
										Next





										'response.write strsql
										If Pdef(t,18) = "NO" Then
											tx = t
										Else
											tx = t - 7
										End If


										PNumCls = Pdef(tx,14)
										If PNumCls = "" Then PNumCls = "Simple"



										If stip <> "H" And stip <> "I"  And stip <> "J" Then
											PNumTip(tx) = Pdef(tx,13)
											If  PNumTip(tx) = "" Then PNumTip(tx) = "1"
										End If

										If stip = "G" Then
											If Pdef(tx,9) = "NO" Then
												MPvin = "NO"
											Else
												MPvin = "SI"
											End If
										End If

										' Anula la numeracion para este tipo o elemento
										If  PNumTip(tx) = "NO" Then PNumCls = PNumTip(tx)

										PFont  = Pdef(tx,0)
										PFsize = Pdef(tx,1)
										PFcol  = Pdef(tx,2)
										Pbgcol = Pdef(tx,3)
										PBorde = Pdef(tx,4)
										PFtipo = Pdef(tx,5)

										If marca = "CA" Then
											PFont  = "VERDANA"
											If smodi Then
												PFcol  = "#000080"
											Else
												PFcol  = "#aaaaaa"
											End If
											PFsize = "8"
											Pbgcol = ""
											PBorde = ""
											PFtipo = ""
										End If


										PSang  = Pdef(tx,6)
										txi = tx
										If tx > 6 Then txi = tx - 7
										PEspa  = Pdef(txi,8)



										Palign = Pdef(tx,16)
										If Palign = "" Then Palign = "left"
										Ptalig = Pdef(tx,17)


										txx = ta

										xmpvin = "NO"

										If stip = "G" And Pdef(txx,9) = "SI" Then
											xmpvin = "SI"
										End If

										If stip = "G" And MPvin  = "NO" Then
											xmpvin = "NO"
										End If

										xPNumS = "NO"

										If stip = "G" Then
											If Pnumer = "SI" Then
												Select Case UCase(Pnumcls)
													Case "SIMPLE"
													xPnumS = "SI"
													Case "COMPUESTA"
													xPnumS = "SI"
												End Select

											End If
										End If


										' TITULOS Y SUBTITULOS --------------------------------------------------

										If stip <> "G" And stip <> "H" And stip <> "I"  And stip <> "J" Then
											'Descripcion
											PNumS = ""
											If Pnumer = "SI" Then
												Select Case UCase(Pnumcls)
													Case "SIMPLE"
													Select Case PnumTip(tx)
														Case "1"
														BnumS = BnumG(bseci)
														Case Else
														BnumS = ConvNumer(BnumG(bseci),PnumTip(tx))
													End Select

													PnumS = BnumS & ". "
													Case "COMPUESTA"
													For i=0 To bseci
														Select Case PnumTip(i + 7)
															Case "NO"
															BnumS = ""
															Case "1"
															BnumS = BnumG(i) & "."
															Case Else
															BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
														End Select

														PnumS = PnumS & BnumS
													Next
													PnumS = PnumS & " "
												End Select

											End If



											'''''''''' arma tabla de indices '''''''''''''''''''''''''



											stexto1 = PnumS & Capitalize(ars(7,inx) & "")
											stexto1 = Replace(stexto1," ","&nbsp;")

											For nn=0 To (bseci+1)
												'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
												stexto1 = "&nbsp;" & stexto1
											Next

											wwsNIV2     = bobj & ""
											wwsNIV3     = (ars(0,inx) & "")

											sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

											xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
											xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
											xverasp = xverasp & "&GLBSOLN=S" & solord(x)
											xverasp = xverasp & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5 '& "&ddd=" & sFuncion


											'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
											iitemC=iitemC+1

											'Se saca border solid
											sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
											iitemC=iitemC+1

											sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
											Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

											'CTABELEN = ofv.GenCelda("","UT  style='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
											CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
											'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")
											wwwpad = "20"
											'if stip <> "A" then wwwpad = "30"
											wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:20px;background-color:" & sikon & ";padding-left:" & wwwpad & "px;'>" &  sLink & "<br></div>" & "<br>"
											'strniv3=strniv3  & "<LI><font color=#fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>"
                                            strniv3=strniv3 & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=Principal href='" & xverasp &  "'> "  &  Capitalize(Replace(STEXTO1,"&nbsp;"," ")) & "</a></div></div></li>"

											'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




										End If



										' PARRAFOS DE TEXTO    --------------------------------------------------

										If stip = "G" Then

											PNumS = ""
											If Pnumer = "SI" Then
												Select Case UCase(Pnumcls)
													Case "NO"
													BnumG(6) = BnumG(6) - 1

													Case "SIMPLE"
													Select Case PnumTip(tx)
														Case "1"
														BnumS = BnumG(6)
														Case Else
														BnumS = ConvNumer(BnumG(6),PnumTip(tx))
													End Select

													PnumS = BnumS & ". "
													Case "COMPUESTA"
													For i=0 To 6
														If BnumG(i) > 0 Then
															Select Case PnumTip(i + 7)
																Case "NO"
																BnumS = ""
																Case "1"
																BnumS = BnumG(i) & "."
																Case Else
																BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
															End Select
															PnumS = PnumS & BnumS
														End If
													Next
													PnumS = PnumS & " "
												End Select

												If UCase(Pnumcls) <> "NO" Then

													sTextoV =  ars(8,inx)

													If Len(sTextoV) > 50 Then sTextoV =  Left(sTextoV,50) & "..."

													'''''''''' arma tabla de indices '''''''''''''''''''''''''



													stexto1 = PnumS & Capitalize(sTextoV & "")
													stexto1 = Replace(stexto1," ","&nbsp;")

													For nn=0 To (bseci+1)
														'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
														stexto1 = "&nbsp;" & stexto1
													Next

													wwsNIV2     = bobj & ""
													wwsNIV3     = (ars(0,inx) & "")

													sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

													'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
													iitemC=iitemC+1

													'Se saca border solid
													sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  &" style='border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
													iitemC=iitemC+1

													sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
													Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

													CTABELEN = ofv.GenCelda("","UT  style='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")

													'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
													wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "ttttt<br><br></div>" & "<br>"
													'strniv3=strniv3  & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 href='"& sAccion & "'>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></LI>"
                                                     strniv3=strniv3 & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=Principal href='" & xverasp &  "'> "  &  Capitalize(Replace(STEXTO1,"&nbsp;"," ")) & "</a></div></div></li>"

													'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

												End If

											End If

										End If

										'------------------------------------------------


										Select Case stip
											Case "A"
											sant     = "A"
											sVinoAnt = sant

											Case "B"
											sant     = "B"
											sVinoAnt = sant
											Case "C"
											sant     = "C"
											sVinoAnt = sant
											Case "D"
											sant     = "D"
											sVinoAnt = sant
											Case "E"
											sant     = "E"
											sVinoAnt = sant
											Case "F"
											sant     = "F"
											sVinoAnt = sant
											Case Else

											sVinoAnt = stip

										End Select

									Next

									If inx > UBound(ars,2) Then
										Exit Do

									End If

									ssale = "no"
									Do Until ssale = "si"

										If     sant = "A" Then

											sant = "X"
										ElseIf sant = "B" Then

											sant = "A"
										ElseIf sant = "C" Then

											sant = "B"
										ElseIf sant = "D" Then

											sant = "C"
										ElseIf sant = "E" Then

											sant = "D"
										ElseIf sant = "F" Then

											sant = "E"
										End If

										If sant = "X" Then
											ssale = "si"

										End If

									Loop

								Loop
								strniv3= strniv3  & "</ul>"
								rsxx1.movenext
							Loop
                             ss11=ss11 & "</ul>"
                              If sFuncion = "E" Then
                               FTABELENUL=FTABELENUL   & "</LI>"
                              else
							   FTABELENUL=FTABELENUL  &  ss11  & strniv3 & "</LI>"
							  end if

						End If
						Call ofv.cerrarconsulta(rsxx1)

						Case 9 ' tablas

						'Strsql = "SELECT *  "
						Strsql = "SELECT * "
						Strsql = Strsql & " from ObjTabDef Where codobjeto = " & sCodigo
						Strsql = Strsql & " order by tabsecuen "
						'  RESPONSE.WRITE STRSQL
						sInput=""

						Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						ss9= "<ul>"
						If   rsx.eof Then


							sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5
						'ss11=ss11  & "<LI class=cn1 data-id='solapa" & solord(x) & wwsNIV2 & "N3' ><font color=  #fdfefe   size=1 face=verdana ><a class=cn1 target=WK1 href='"& sAccion & "'>" & UCase(Nombre_Solapa) & "</a></font>"
							FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text'  data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=WK1 href='" & sAccion &  "'>"  &  UCase(Nombre_Solapa)  & "</a></div></div></li>"

						Else
					   	 	sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5
							xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="

					  		 xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
					   		xverasp = xverasp & "&GLBSOLN=S" & solord(x)
			           		xverasp = xverasp & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5 '& "&ddd=" & sFuncion

					    If sFuncion = "E" Then
						  FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a class=ntmnn target=Principal data-id='s1solapa" & solord(x) & "' class=""link"" href='"&xverasp & "'>"  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"
					    else
						   FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a class=ntmnn target=Principal data-id='s1solapa" & solord(x) & "' class=""link"" xhref='"&xverasp & "'>"  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"
					    end if

							'FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  Capitalize(Nombre_Solapa)  & "</a></div></div>"

							Do While Not rsx.eof


								'sfidkD  = sfidkB
								stexto1 = Capitalize(rsx(4) & "")
								stexto1 =rsx(5) & "." &  Replace(stexto1," ","&nbsp;")

								Texto2 = rsx(2)
								wwsNIV2     = (rsx(0) & "")
								wwsNIV3     = "0"
								sAccion =xSolaspd & solord(x) & "&NIV1=" & solord(x)  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5

                                xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
								xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
								xverasp = xverasp & "&GLBSOLN=S" & solord(x)
								xverasp = xverasp & "&NIV1=" & solord(x)   & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5 '& "&ddd=" & sFuncion

						        ss9=ss9  & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a  class=ntmnn target=Principal href='" & xverasp &  "'> "  &  Capitalize(Replace(STEXTO1,"&nbsp;"," ")) & "</a></div></div></li>"


								rsx.movenext




							Loop
						End If
						'	FTABELENUL=FTABELENUL & "</UL>"
						ss9=ss9 & "</UL>"
						If sFuncion = "E" Then
						 FTABELENUL=FTABELENUL   & "</LI>"
						 else
						   FTABELENUL=FTABELENUL    & ss9 & "</LI>"
					   end if


						Call ofv.cerrarconsulta(rsx)
						Case Else
						sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5
                        xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="

						xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
						xverasp = xverasp & "&GLBSOLN=S" & solord(x)
						xverasp = xverasp & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=0" & wwsNIV4  & "&NIV5=0" & wwsNIV5 '& "&ddd=" & sFuncion

						FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(Nombre_Solapa) & "'><a class=ntmnn target=Principal HREF='"& xverasp & "'> "  &  Capitalize(Nombre_Solapa) & "</a></div></div></li>"
					End Select
					'FTABELENUL=FTABELENUL  &"</LI>"
					sinput2 = ""

					If sFuncion = "E" Or sFuncion = "ED"  Then

					Else
						'             xxxx = GetIndice(solord(x))
						'     sinput2 = ""
						sinput2 = sinput2 & "<div id='SOLX" & solord(x) & "' style='xdisplay:none; border-style:solid;border-width:1;border-color:#ff3333;position: absolute;top:1;left:1;' xxstyle='font-weight:bold; font-size:12px; background-color:" & sikon & " ; border-radius: 60px; width: 167px; height: 40px; border-style:solid;border-width:1px;border-color:grey;' >"
						sinput2 = sinput2' & GetIndice(solord(x))
						sinput2 = sinput2 & "</div>"

					End If


					CTABELEN = ofv.GenCelda("",  "uT align='center'" & wwwestilo  & " ","","","","","", sInput2  ,"si")
					FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
					'FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"
					'GetIndice(solord(x))
					'"uT align='center'"
					'& wwwestilo

					wwwk1 = wwwk1 & "<div  style='xdisplay:none;width=200px;height=200px;display: inline-block;margin: 30px 30px 30px 30px;' >" & sLink & "</div>"



				Else

					t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
					t3 = t3 & sInput
					t3 = t3 & "</A>"

					t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

					r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

				End If

				'response.write clng(sSClick) & " solact "  & clng(solord(x)) & chr(13)
				If CLng(sSClick) = CLng(solord(x)) Then
					session("codSol")=sSClick
					' response.write clng(sSClick)
					xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
					xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
					xverasp = xverasp & "&GLBSOLN=S" & solord(x)
					xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

					sHTML1 = ""

					sHTML1 = sHTML1 & "<form name=SSOLACT id=SSOLACT  action='" & xverasp & "' "
					sHTML1 = sHTML1 & " method=post target='"
					If CLng(sSClick) = 12 Then
						sHTML1 = sHTML1 & "_blank"
					Else
						sHTML1 = sHTML1 & "Principal"
					End If
					sHTML1 = sHTML1 & "' style='display:none;' >" & Chr(13)
					sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & Chr(13)
					sHTML1 = sHTML1 & "</form>" & Chr(13)

					'''''''''''''''''''''''''''''''''
					' NUEVO PARA VERSION 4.6

					'response.write request.querystring("Test") & " pregunta por Test " & "<br>"

					If request.querystring("Test") = "T" Then

						sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
						sHTML1 = sHTML1 & " window.top.CorSol='" & Replace(UCase(sSolapa),"*"," ") & "'; " & Chr(13)
						sHTML1 = sHTML1 & "</script>" & Chr(13)

					Else
						sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
						sHTML1 = sHTML1 & " window.top.CorSol='" & Replace(UCase(sSolapa),"*"," ") & "'; " & Chr(13)

						sHTML1 = sHTML1 & " var element =  document.getElementById(window.top.imenu);"

						sHTML1 = sHTML1 & " if (typeof(element) != 'undefined' && element != null)"
						sHTML1 = sHTML1 & " {"

						''''''''Agregado 10032015 - Version 4.6
						sHTML1 = sHTML1 & " element.style.backgroundColor='"&session("CMICTT")&"';"
						sHTML1 = sHTML1 & " element.style.color='"&session("CMICTTFT")&"';"

						''''''''Agregado 10032015 - FIN - Version 4.6

						sHTML1 = sHTML1 & " }"

						sHTML1 = sHTML1 & "</script>" & Chr(13)

					End If
					' NUEVO PARA VERSION 4.6 -FIN-
					'''''''''''''''''''''''''''''''''

					If session("WEBDESK") = "S" Then




					Else

						If  Session("FULLGUI") = "N" Then
							styl = "background-color:#888888;"
						Else
							styl = XspEx("#888888","#cccccc",0) & ";"
						End If

						xlfun = " onclick=""document.getElementById('MDBEDL').style.display=''; document.getElementById('DBEDTD1').style.display='none'; document.getElementById('DBEDTD2').style.display=''; document.getElementById('MM2').style.display=''; "" "
						xlfun2 = " onclick=""document.getElementById('MDBEDL').style.display='none'; document.getElementById('DBEDTD1').style.display=''; document.getElementById('DBEDTD2').style.display='none'; document.getElementById('MM2').style.display='none'; "" "

						macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:10px;' " & xlfun,"","90%","","","",Replace(UCase(sSolapa),"*"," "),"si")

						xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


						macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
						macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

						marow = ofv.GenRow("","","","",lengtd,"","",macel,"si")

						sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

					End If

				End If


			End If
		Next

		wwwk1 = wwwk1 & "</div>"

		CTABELEN = ofv.GenCelda("","uT align='center' ","","","","","",wwwk1,"si")
		FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
		'FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"

		If SMAINAPL = "OR" Or SMAINAPL = "PL" Then

		Else

			''-----------------------------------------------------------------------------------------------------

			''''''''''''''''''''''''''''''''''''''''''''
			'' Accesos directos del documento ''''''''''
			''''''''''''''''''''''''''''''''''''''''''''

			Dim ATABSCPT(200,2)
			Dim bConsultar()
			ISCPT = 0
			TSCPT = 0

			''    response.write "entro1"

			If sFuncion <> "X"   Then

				''  response.write "entro2"

				ReDim Preserve bConsultar(4)

				If sFuncion = "E"   Then

					'Se comenta para que no se muestren las normas relacionadas

					'bConsultar(2) = 0
					'bConsultar(4) = 0

					'bConsultar(1) = 1
					'bConsultar(3) = 1

					'else

					'bConsultar(2) = 0
					'bConsultar(4) = 0

					'bConsultar(1) = 0
					'bConsultar(3) = 1

				End If


				sVerEstado = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

				sverestado = sverestado & " and Objprty.Estado In (9,"
				For i = 1 To 4
					If bConsultar(i) Then sVerEstado = sVerEstado & i & ","
				Next
				sVerEstado = Left(sVerEstado, Len(sVerEstado) - 1) & ") "

				'''''''''''''''''''''''''''''''''''''''''



				sActcon  = "Consultar.asp?Codigo=NNNN&Tipo=TTTT"
				sActcon  = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
				sHref    = "&nbsp;"
				' response.write sActcon
				sFuncion = " onclick=" & Chr(34) & " window.top.abreWin('" & sActcon & "','TTTT'); return false;" & Chr(34)


				Strsql = "SELECT * FROM objacd  "
				Strsql = Strsql & " Where codobjeto=" & sCodigo
				Strsql = Strsql & " And tipobjeto = '" & stipo & "' "

				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				Do Until rs.eof

					''   response.write "entro3"


					If  rs(3) = "S" Then
						'''' Carpeta superior

						If  rs(6) = "S" Then


							''----------------

							strsql  = "select * from csbarr where codojasoc = " & sCodigo
							Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
							Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If  Not rs.eof Then
								essCodigo = rs1(1)
								essTipo   = rs1(2)

							Else
								essCodigo = 0
								essTipo   = "NO"
							End If
							Call ofv.cerrarconsulta(rs1)


							sNIVE   = "no"

							StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
							StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
							Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If Not rsIcono.EOF Then

								sIcono  = rsIcono(0)
								sTipoE  = rsIcono(1)
								sPerfil = rsIcono(2)
								sNIVE   = rsIcono(3)

							End If
							Call OFV.cerrarconsulta(rsIcono)

							If sNIVE   = "C"  Then

								strsql  = "select * from csbarr where codojasoc = " & essCodigo
								Strsql  = Strsql & " and tipobjasoc = '" & essTipo & "'"
								Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								If  Not rs.eof Then

									essCodigo = rs1(1)
									essTipo   = rs1(2)

								End If
								Call ofv.cerrarconsulta(rs1)


								'''''''''''''''''''''''''''''''''''''''''

								StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
								StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
								StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
								StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
								StrSql = StrSql & " where "
								StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
								StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
								StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
								StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
								StrSql = StrSql & "  objetos.objun =  1 )  and "
								StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
								StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
								StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
								''  response.write "<br>" & strsql & "<br>"
								sVino = "N"
								Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
								If  Not rs2.EOF Then

									Do Until rs2.EOF

										If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

											sVinook = "N"
											If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
												'' response.write "<br><br>entro"
												sVino   = "S"
												sVinook = "S"
											End If

											If SESSION("EMPMUN") = "S" Then
												StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
												StrSql = StrSql & " where "
												StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
												StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
												Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If  Not rsmas2.EOF Then
													sVino   = "S"
													sVinook = "S"
												End If
											End If

											sNIVE   = "D"

											If  sVinook = "S" Then

												stitem = rs2(4)
												StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
												StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
												Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If Not rsIcono.EOF Then
													sIcono  = rsIcono(0)
													sTipoE  = rsIcono(1)
													sPerfil = rsIcono(2)
													sNIVE   = rsIcono(3)

													sitem = sTipoE & ": " & rs2(1)

												End If
												Call OFV.cerrarconsulta(rsIcono)

												If sNIVE   = "D"  Then


													sSolapa =  sitem

													If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

														sInput = Replace(UCase(sSolapa),"*"," ")

													Else

														sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"

													End If

													sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
													sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))





													If session("WEBDESK") = "S" Then


														'sInput  = ">" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
														'sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
														'Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

														'wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

														'CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
														'FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")




													Else

														t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
														t3 = t3 & sInput
														t3 = t3 & "</A>"

														t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

														r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

													End If





												Else





												End If


											End If

										End If

										rs2.movenext
									Loop


								End If



								'''''''''''''''''''''''''''''''''''''''

							End If

							''----------------



						End If





					ElseIf  rs(3) = "P" Then
						'''' Carpeta primaria

						If  rs(6) = "S" Then

							strsql  = "select * from csbarr where codojasoc = " & sCodigo
							Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
							Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If  Not rs.eof Then
								essCodigo = rs1(1)
								essTipo   = rs1(2)

							Else
								essCodigo = 0
								essTipo   = "NO"
							End If
							Call ofv.cerrarconsulta(rs1)


							sNIVE   = "no"

							StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
							StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
							Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If Not rsIcono.EOF Then

								sIcono  = rsIcono(0)
								sTipoE  = rsIcono(1)
								sPerfil = rsIcono(2)
								sNIVE   = rsIcono(3)

							End If
							Call OFV.cerrarconsulta(rsIcono)




							StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
							StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
							StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
							StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
							StrSql = StrSql & " where "
							StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
							StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
							StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
							StrSql = StrSql & " (objetos.objun =  0" & session("UN") & " or "
							StrSql = StrSql & "  objetos.objun =  1 )  and "
							StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
							StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
							StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
							''   response.write "<br>" & strsql & "<br>"
							sVino = "N"
							Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If  Not rs2.EOF Then

								Do Until rs2.EOF

									If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

										sVinook = "N"
										If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
											'' response.write "<br><br>entro"
											sVino   = "S"
											sVinook = "S"
										End If

										If SESSION("EMPMUN") = "S" Then
											StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
											StrSql = StrSql & " where "
											StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
											StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
											Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If  Not rsmas2.EOF Then
												sVino   = "S"
												sVinook = "S"
											End If
										End If

										sNIVE   = "D"

										If  sVinook = "S" Then

											stitem = rs2(4)
											StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
											StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
											Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If Not rsIcono.EOF Then
												sIcono  = rsIcono(0)
												sTipoE  = rsIcono(1)
												sPerfil = rsIcono(2)
												sNIVE   = rsIcono(3)

												sitem = sTipoE & ": " & rs2(1)

											End If
											Call OFV.cerrarconsulta(rsIcono)

											If sNIVE   = "D"  Then


												sSolapa =  sitem

												If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
													sInput = Replace(UCase(sSolapa),"*"," ")
												Else
													sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
												End If

												sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
												sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))




												If session("WEBDESK") = "S" Then




													'    sInput  = "" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
													'    sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:normal;'  " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
													'    Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

													'    CTABELEN = ofv.GenCelda("","UT  style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' ","","","","","",sLink,"si")
													'    FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



													'' sInput  = ">" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
													sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
													Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

													wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

													CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
													FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
													'	FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"




												Else

													t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
													t3 = t3 & sInput
													t3 = t3 & "</A>"

													t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

													r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

												End If


											Else


											End If


										End If

									End If

									rs2.movenext
								Loop


								'''''''''''''''''''''''''''''''''''''''

							End If




						End If




					ElseIf  rs(3) = "R" Then
						'''' Carpeta dependiente

						If  rs(6) = "S" Then

							''----------------

							TSCPT = TSCPT + 1

							ATABSCPT(TSCPT,0) = ""
							ATABSCPT(TSCPT,1) = sCodigo
							ATABSCPT(TSCPT,2) = sTipo



							strsql  = "select * from csbarr where codojasoc = " & sCodigo
							Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
							Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If  Not rs.eof Then
								essCodigo = rs1(1)
								essTipo   = rs1(2)

							Else
								essCodigo = 0
								essTipo   = "NO"
							End If
							Call ofv.cerrarconsulta(rs1)


							sNIVE   = "no"

							StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
							StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
							Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If Not rsIcono.EOF Then

								sIcono  = rsIcono(0)
								sTipoE  = rsIcono(1)
								sPerfil = rsIcono(2)
								sNIVE   = rsIcono(3)

							End If
							Call OFV.cerrarconsulta(rsIcono)

							If sNIVE   = "C"  Then

								'  essCodigo = sCodigo
								'  essTipo   = sTipo

								'''''''''''''''''''''''''''''''''''''''''

								StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
								StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
								StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
								StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
								StrSql = StrSql & " where "
								StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
								StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
								StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
								StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
								StrSql = StrSql & "  objetos.objun =  1 )  and "
								StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
								StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
								StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
								''   response.write "<br>" & strsql & "<br>"
								sVino = "N"
								Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
								If  Not rs2.EOF Then

									Do Until rs2.EOF

										If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

											sVinook = "N"
											If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
												'' response.write "<br><br>entro"
												sVino   = "S"
												sVinook = "S"
											End If

											If SESSION("EMPMUN") = "S" Then
												StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
												StrSql = StrSql & " where "
												StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
												StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
												Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If  Not rsmas2.EOF Then
													sVino   = "S"
													sVinook = "S"
												End If
											End If

											sNIVE   = "D"

											If  sVinook = "S" Then

												stitem = rs2(4)
												StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
												StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
												Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
												If Not rsIcono.EOF Then
													sIcono  = rsIcono(0)
													sTipoE  = rsIcono(1)
													sPerfil = rsIcono(2)
													sNIVE   = rsIcono(3)

													sitem = sTipoE & ": " & rs2(1)

												End If
												Call OFV.cerrarconsulta(rsIcono)

												If sNIVE   = "D"  Then




												Else


													TSCPT = TSCPT + 1

													ATABSCPT(TSCPT,0) = ""
													ATABSCPT(TSCPT,1) = rs2(0)
													ATABSCPT(TSCPT,2) = rs2(2)


												End If


											End If

										End If

										rs2.movenext
									Loop

								End If



								Do While Not ISCPT > TSCPT

									If ISCPT = 0 Then ISCPT = 1


									For ISCPT=ISCPT To TSCPT



										StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
										StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
										StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
										StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
										StrSql = StrSql & " where "
										StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
										StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
										StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
										StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
										StrSql = StrSql & "  objetos.objun =  1 )  and "
										StrSql = StrSql & " RELAOBJ.codobjmaster = 0" & ATABSCPT(ISCPT,1)
										StrSql = StrSql & " and relaobj.tipobjmaster = '" & ATABSCPT(ISCPT,2) & "' " & sVerEstado
										StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
										''    response.write "<br>" & strsql & "<br>"
										sVino = "N"
										Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
										If  Not rs2.EOF Then

											Do Until rs2.EOF

												If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then



													sVinook = "N"
													If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
														'' response.write "<br><br>entro"
														sVino   = "S"
														sVinook = "S"
													End If

													If SESSION("EMPMUN") = "S" Then
														StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
														StrSql = StrSql & " where "
														StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
														StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
														Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
														If  Not rsmas2.EOF Then
															sVino   = "S"
															sVinook = "S"
														End If
													End If

													sNIVE   = "D"

													If  sVinook = "S" Then

														stitem = rs2(4)
														StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
														StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
														Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
														If Not rsIcono.EOF Then
															sIcono  = rsIcono(0)
															sTipoE  = rsIcono(1)
															sPerfil = rsIcono(2)
															sNIVE   = rsIcono(3)

															sitem = sTipoE & ": " & rs2(1)

														End If
														Call OFV.cerrarconsulta(rsIcono)

														If sNIVE   = "D" Then


															sSolapa =  sitem

															If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
																sInput = Replace(UCase(sSolapa),"*"," ")
															Else
																sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
															End If

															sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
															sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))



															If session("WEBDESK") = "S" Then



																'  sInput  = "" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
																'  sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:normal;'  " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
																'  Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

																'  CTABELEN = ofv.GenCelda("","UT  style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' ","","","","","",sLink,"si")
																'  FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")


																sInput  = ">" & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;") & ""
																sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
																Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

																wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

																CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
																FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
																' FTABELENUL=FTABELENUL & "<LI>"  &  Nombre_Solapa & "</LI>"

															Else

																t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
																t3 = t3 & sInput
																t3 = t3 & "</A>"

																t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

																r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

															End If



														Else



															TSCPT = TSCPT + 1

															ATABSCPT(TSCPT,0) = ""
															ATABSCPT(TSCPT,1) = rs2(0)
															ATABSCPT(TSCPT,2) = rs2(2)


														End If


													End If


												End If

												rs2.movenext
											Loop

										End If


									Next

								Loop



							End If



						End If


					Else

						''''' documentos o subcarpetas independientes


					End If


					''''''''''''''''''''''''''''


					'''''''''''''''''''''''''''''



					rs.movenext
				Loop

			End If

			''''''''''''''''''''''''''''''''''''''''''''


			''-----------------------------------------------------------------------------------------------------


		End If



		If session("WEBDESK") = "S" Then


			'''    sHTML = sHTML & Celda
			'''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

			'sHTML = sHTML & "<div id='WKTBSP1'  >"
			'FTABELENUL=FTABELENUL & "</UL></NAV>" & ss11  & ss7 & ss9 &"</DIV>"
			'FTABELENUL=FTABELENUL & "</UL></NAV>" & ss11  & ss7 & ss9 &"</DIV>"
			'	sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "aaa  ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN , "si")
			'	shtml = sHTML &  FTABELENUL  & "</div>"

		Else

			'	sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"


			If  Session("FULLGUI") = "N" Then
				styl = "background-color:#888888;"
			Else
				styl = XspEx("#888888","#cccccc",0) & ";"
			End If

			styl = ""

			If ofv.ModTitCls = "1" Then
				'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			Else
				'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			End If

			sHTML = sHTML & "</div>"

		End If








		sHTML = sHTML & sHTML1
		sHTML = sHTML & sinput2


		Call ofv.cerrarconn(cn)

		sHTML = sHTML & "</body></html>"


	End If
	'If session("APLMDCLI")="BGAL" Then
	'Response.Write Replace(sHTML,"&nbsp;"," ") 'sHTML
	sHTML=""
	Whtml1 = ""
	Whtml1 = Whtml1 & "<script language=javascript >"
	'				Whtml1 = Whtml1 & " ClosIndice();" & Chr(13)

	Whtml1 = Whtml1 & "</script>"

	armasolapa=FTABELENUL & Whtml1

	'Else
	'	Response.Write SHTML
	'End If

End Function

Function GetIndice33(gtsolapa)

	FTABELENX = ""

	wwsNIV1     = gtsolapa
	wwsNIV2     = "0"
	wwsNIV3     = "0"
	wwsNIV4     = "0"
	wwsNIV5     = "0"


	sft   = "<font face=tahoma,arial size=1 color=#336699>&nbsp;<b>"
	sfta  = "<font face=tahoma,arial size=1 color=#993333>&nbsp;<b>"
	sfidc = "<font face=corflaticon size=2 color=#336699>0</font>"
	sfida = "<font face=corflaticon size=2 color=#993333>1</font>"
	sfitc = "<font face=corflaticon size=2 color=#336699>0</font>"
	sfita = "<font face=corflaticon size=2 color=#993333>1</font>"
	sftc  = "</b></font>"


	'  ASPs(1)  = "edPropiedades.asp?SKIND=1"
	'  ASPs(2)  = "CONWEDT.asp?SKIND=1"
	'  ASPs(3)  = "CoSectores.asp?SKIND=1"
	'  ASPs(4)  = "edEventos.asp?SKIND=1"
	'  ASPs(5)  = "CoProcedimientos.asp?SKIND=1"
	'  ASPs(6)  = "CoObjDsc.asp?SKIND=1"
	'  ASPs(7)  = "CoControles.asp?SKIND=1"
	'  ASPs(8)  = "CoEsquemas.asp?SKIND=1"
	'  ASPs(9)  = "CoTablasCons.asp?SKIND=1"
	'  ASPs(10) = "CoArchCons.asp?SKIND=1"
	'  ASPs(11) = "CoestrCons.asp?SKIND=1"
	'  ASPs(12) = "CursogramaLnch.asp?SKIND=1"
	'  ASPs(13) = "CoArchCons.asp?SKIND=2"
	'  ASPs(14) = "CoFAQ.asp?SKIND=1"
	'  ASPs(15) = "CoGeneral.asp?SKIND=1"
	'  ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"

	gtselinx = ""

	wwwbloke = ""
	Select Case CStr(gtsolapa & "")



		Case "1"

		Case "2"

		Case "3"

		Case "4"

		Case "5"

		eveant = ""

		Strsql = "SELECT *  "
		Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & sCodigo
		''  Strsql = Strsql & sSelecx
		'' strsql = strsql & addperfil
		strsql = strsql & " order by secuen, secuencia,codproced "
		'response.write strsql
		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		Do While Not rsx.eof

			If eveant <> rsx(4) Then

				'    Celdat  = "<td class=uts >" & sfidkA & "</td>"
				'    Celdat  = Celdat & "<td class=utg  >" & rsx(2) & ".&nbsp;" & rsx(4) & "</td>"
				'    Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

				eveant  = rsx(4)
			End If

			sAccion = "CoProcedimientos.asp?INXMOD=N&CantRegAMover=A&Codobjeto=" & sCodigo & "&Menu=" & sMenux
			sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&proced=" & rsx(0) & "&test=" & sTest
			sOnClik1 = ""
			sOnClik2 = " onclick=""document.getElementById('ASC" & rsx(0) & "').click();"" "

			'  if rsx(0)  = rs(0).value then
			'     sfidkD  = sfidkc
			'     stexto1 = "<u>" & rsx(6) & ".&nbsp;" & rsx(3) & "&nbsp;-&nbsp;" & rsx(5) & "</u>"

			'  else
			sfidkD  = sfidkB
			stexto1 = rsx(2) & "." & rsx(6) & "."  & Capitalize(rsx(4) & "")  & "." & Capitalize(rsx(3) & "") & " - " & Capitalize(rsx(5) & "")
			stexto1 = Replace(stexto1," ","&nbsp;")

			Texto2 = rsx(2)

			'  end if

			'  Celdat  = stexto1

			'  stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=nt href='" & sAccion & "' " & sOnClik1 & " "
			'  stexto1 = stexto1 & "   >" & Celdat & "</a>"

			'  Celdat  = "<td class=utis >" & sfidkD & "</td>"
			'  Celdat  = Celdat & "<td class=uti style='cursor:pointer;' " & sOnClik2 & " >" & stexto1 & "</td>"
			'  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

			wwsNIV2     = (rsx(0) & "")
			wwsNIV3     = "0"

			sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5



			'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >aaaaaaaa" & stexto1 & "</font>"
			iitemC=iitemC +1

			'Se saca border solid
			sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC &" style='margin:2px;' >" & stexto1 & "</font>"
			iitemC=iitemC +1

			'border-radius:20px;
			'border-width:2;
			'border-color:black;
			sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput , "si")


			Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp</span> "




			'CTABELEN = ofv.GenCelda("","UT  style='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink ,"si")


			CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink ,"si")

			' FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")

			wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "<br><br></div>" & "<br>"




			rsx.movenext




		Loop

		wwwdisplay = "none"

		If (sSClick & "") = (gtsolapa & "") Then wwwdisplay = "block"

		'Se le saca color a procedimientos & sSClick y
		'background-color:" & sikon & ";

		'Se saca border solid y se agrega radius
		wwwbloke1 = "<div class='BLK' id='BLOK" & gtsolapa & "' style='display:" & wwwdisplay & ";' >" & wwwbloke & "</div>"

		CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",wwwbloke1 ,"si")

		FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

		Case "6"

		Case "7" '' controles


		wwbbusca = "0"

		Strsql = "SELECT * from cscontrolsel "
		Strsql = Strsql & " Where codobjeto = " & sCodigo
		Strsql = Strsql & " and seldes is NOT NULL and descri is NOT NULL "
		'''''''Strsql = Strsql & " and descri is NOT NULL "

		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		'contar = 0
		Do While Not rsx.eof

			'contar = contar + 1
			'response.write contar & "contar " & "<br>"

			If CStr(rsx(0) & "")   = wwbbusca Then
				sfidkD  = sfidkC
				stexto1 = rsx(1)
				stextoact = stexto1
			Else
				sfidkD   = sfidkA
				stexto1  = rsx(1) & ""
				''
			End If
			''
			sAccion  = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
			sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & rsx(0) & "&test=" & sTest
			sOnClik1 = ""
			sOnClik2 = " onclick=""document.getElementById('ASC" & rsx(0) & "').click();"" "

			'    Celdat  = stexto1
			'    stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
			'    stexto1 = stexto1 & "   >" & Celdat & "</a>"

			'     Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
			'     Celdat  = Celdat & "<td class=ut  >" & stexto1  & "zz" & "</td>"
			'     Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

			stexto1 = Replace(Capitalize(stexto1 & "")," ","&nbsp;")

			wwsNIV2     = CStr(rsx(0) & "")
			wwsNIV3     = "0"

			sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

			'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
			iitemC=iitemC +1

			'Se saca border solid
			sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
			iitemC=iitemC +1

			sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
			Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

			'CTABELEN = ofv.GenCelda("","UT  style='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
			CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
			' FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")

			wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "<br><br></div>" & "<br>"


			''
			'end if
			''

			rsx.movenext
		Loop

		wwwdisplay = "none"

		If (sSClick & "") = (gtsolapa & "") Then wwwdisplay = "block"

		'Se saca border solid y size y se agrega radius y sSClick & " " &
		wwwbloke1 = "<div class='BLK' id='BLOK" & gtsolapa & "' style='display:" & wwwdisplay & ";border-radius:10px;' >" &  wwwbloke & "</div>"

		CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",wwwbloke1 ,"si")

		FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

		Case "8" '' esquemas

		wwbbusca = "0"

		Strsql = "SELECT * from csesquemasel "
		Strsql = Strsql & " Where codobjeto = " & sCodigo

		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		Do While Not rsx.eof

			If CStr(rsx(0) & "")   = wwbbusca Then
				sfidkD  = sfidkC
				stexto1 = rsx(1)
				stextoact = stexto1
			Else
				sfidkD   = sfidkA
				stexto1  = rsx(1) & ""
				sAccion  = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & sCodigo & "&Menu=" & sMenux
				sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & rsx(0) & "&test=" & sTest
				sOnClik1 = ""
				sOnClik2 = " onclick=""document.getElementById(ASC" & rsx(0) & "').click();"" "

				'    Celdat  = stexto1
				'    stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
				'    stexto1 = stexto1 & "   >" & Celdat & "</a>"

				'     Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
				'     Celdat  = Celdat & "<td class=ut  >" & stexto1  & "zz" & "</td>"
				'     Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

				stexto1 = Replace(Capitalize(stexto1 & "")," ","&nbsp;")

				wwsNIV2     = CStr(rsx(0) & "")
				wwsNIV3     = "0"

				sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

				sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
				iitemC=iitemC+1

				'Se saca border solid
				sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
				iitemC=iitemC+1


				sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
				Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				'CTABELEN = ofv.GenCelda("","UT  style='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
				CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
				'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")

				wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "<br><br></div>" & "<br>"

			End If

			rsx.movenext
		Loop

		wwwdisplay = "none"

		If (sSClick & "") = (gtsolapa & "") Then wwwdisplay = "block"

		'Se saca border solid y size y se agrega radius y & sSClick & " "
		wwwbloke1 = "<div class='BLK' id='BLOK" & gtsolapa & "' style='display:" & wwwdisplay & ";' >" & wwwbloke & "</div>"

		CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",wwwbloke1 ,"si")

		FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")


		Case "9"  ''' bloque de tablas

		Strsql = "SELECT * "
		Strsql = Strsql & " from ObjTabDef Where codobjeto = " & sCodigo
		Strsql = Strsql & " order by tabsecuen "

		Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsxx1.eof Then
			'' if sdocum = "" then sdocum = cstr(rsxx1(0))

			' tomo todas las tablas

			stexto1 = "Seleccionar&nbsp;tabla"
			'  Celdat  = stexto1

			sAccion = "javascript:void(0);"
			sOnClik = " onclick=""document.getElementById('TIND2').style.display='';document.getElementById('TIND1').style.display='none';"" "


			stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
			stexto1 = stexto1 & "   >" & Celdat & "</a>"


			'  Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
			'  Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"


			stexto1 = "Tablas"

			'  Celdat  = "<td class=uts >" & sfidk0 & "</td>"
			' Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
			' Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")


			Do Until rsxx1.eof

				wwsarch  = "0" & rsxx1(0)



				'    if  sdocum = cstr(rsxx1(0)) then
				'        sarch  = rsxx1(0)
				'    end if

				'   if  rsxx1(3) <> "" then
				''  Celda  = ofv.GenCelda("","cc","","1%","","","",stexto1,"si")
				''  Celda  = Celda & ofv.GenCelda("","","","2","","","","&nbsp;","si")
				''  Formulario = Formulario & Celda
				'  end if


				'   if  sdocum = cstr(rsxx1(0)) then
				'       sfidkD  = sfidkC
				'       stexto1 = replace(rsxx1(4)," ","&nbsp;")
				'       stextoact = stexto1

				'   else
				sfidkD  = sfidkA
				''''    stexto1 = replace(rsxx1(4)," ","&nbsp;")

				stexto1 = rsxx1(5) & "." & Replace(Capitalize(rsxx1(4) & "")," ","&nbsp;")
				sAccion = "CoTablasCons.asp?INXMOD=N&Codobjeto=" & scodigo & "&Tipobjeto=" & stipo & "&sdocum=" & rsxx1(0)
				sOnClik1 = ""
				sOnClik2 = " onclick=""document.getElementById('ASC" & rsxx1(0) & "').click();"" "


				Celdat  = stexto1

				'  stexto1 = "<A ID='ASC" & rsxx1(0) & "' NAME='ASC" & rsxx1(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
				'  stexto1 = stexto1 & "   >" & Celdat & "</a>"

				'  Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
				'  Celdat  = Celdat & "<td class=uti  >" & stexto1 & "</td>"
				'  Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

				wwsNIV2     = (rsxx1(0) & "")
				wwsNIV3     = "0"

				sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

				'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
				iitemC=iitemC+1

				'Se saca border solid
				sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC &" style='border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
				iitemC=iitemC+1

				sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "no")
				Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				'CTABELEN = ofv.GenCelda("","UT  style='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
				CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
				'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")

				wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "<br><br></div>" & "<br>"

				'  end if

				Strsql = "SELECT * From Objtabblk Where Tabdef = " & wwsarch & " order by tabblksec"
				Set rsxx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				Do Until rsxx2.eof
					wwnID      = rsxx2(0)
					wwnSec     = rsxx2(4)
					wwnBTitulo = rsxx2(5)

					If Len(wwnBTitulo) > 0 Then
						wwltitu = Len(wwnBTitulo)
						If  wwnBTitulo <> Space(wwltitu) And UCase(wwnBTitulo) <> "NO" Then

							stexto1 = wwnSec & "." & Replace(Capitalize(wwnBTitulo & "")," ","&nbsp;")
							stexto1 = Replace(stexto1," ","&nbsp;")

							'  Celdat  = stexto1
							'  stexto1 = "<A class=ut href='javascript:void(0);' "
							'  stexto1 = stexto1 & "   >" & Celdat & "</a>"
							'  Celdat  = "<td class=utis style='cursor:pointer;' onclick=""document.all['TBLMAIN'].style.visibility='hidden';creawin('" & sPoint & "');document.all['TBLMAIN'].style.visibility='visible';"" >" & sfidkB & "</td>"
							'  Celdat  = Celdat & "<td class=uti onclick=""document.all['TBLMAIN'].style.visibility='hidden';creawin('" & sPoint & "');document.all['TBLMAIN'].style.visibility='visible';"" >" & stexto1 & "</td>"
							'  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

							wwsNIV2     = (rsxx1(0) & "")
							wwsNIV3     = (wwnID & "")

							sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

							sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
							iitemC=iitemC+1

							'Se saca border solid
							sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-width:1;border-bottom-color:#bbbbbb;padding-left:10px;' >" & stexto1 & "</font>"
							iitemC=iitemC+1

							sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
							Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

							'CTABELEN = ofv.GenCelda("","UT  style='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:40px;' ","","","","","",sLink,"si")
							CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:40px;' ","","","","","",sLink,"si")
							'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")

							wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:5px;background-color:" & sikon & ";'>" &  sLink & "<br></div>" & "<br>"

						End If
					End If

					rsxx2.movenext
				Loop



				rsxx1.movenext
			Loop
			wwwdisplay = "none"

			If (sSClick & "") = (gtsolapa & "") Then wwwdisplay = "block"

			'Se saca border solid y size y se agrega radius y & sSClick & " "
			wwwbloke1 = "<div class='BLK' id='BLOK" & gtsolapa & "' style='display:" & wwwdisplay & ";' >" & wwwbloke & "</div>"

			CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",wwwbloke1 ,"si")

			FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

		End If


		Case "10"

		Case "11" '' bloque de contenido ----------------------------------------------------------------------------------------

		Dim Inx
		Dim nrs(100)

		ReDim Preserve bsec(8)
		ReDim Preserve bnum(8)
		ReDim Preserve bnumG(8)


		Dim Pdef(13,18)

		strsql = "Select * From Objestr Where codobjeto =" & sCodigo
		Strsql = Strsql & " order by CONsecuen "
		'  response.write strsql & "<br>"
		Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		bsecuen = ""
		bDescri = " "
		bobj = ""
		Do While Not rsxx1.eof
			bsecuen = rsxx1(4)
			bDescri = rsxx1(3)
			bobj    = rsxx1(0)

			'  stexto1 = "<A ID='ATIT" & rsxx1(0) & "' NAME='ATIT" & rsxx1(0) & "' "  & chr(13)
			'  stexto1 = stexto1 & " class=nt href='coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
			'  stexto1 = stexto1 & "&sdocum=" & sdocum & "&sdocc=" & rsxx1(0)
			'  stexto1 = stexto1 & "'   >" & sfidkc & "&nbsp;<big>" & rsxx1(3) & "</big></a>" & chr(13)

			'  Celdat = "<td class=ut style='cursor:pointer;'  onclick=""document.all.ATIT" & rsxx1(0) & ".click();this.style.backgroundColor='red';"" >" & stexto1 & "</td>"
			'  Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si") & chr(13)

			stexto1 = bsecuen & "." & Replace(Capitalize(bDescri & "")," ","&nbsp;")
			stexto1 = Replace(stexto1," ","&nbsp;")

			wwsNIV2     = bobj & ""
			wwsNIV3     = "0"

			sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

			'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & " </font>"
			iitemC=iitemC+1

			'Se saca border solid
			sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & " </font>"
			iitemC=iitemC+1

			sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
			Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

			'CTABELEN = ofv.GenCelda("","UT  style='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:10px;' ","","","","","",sLink,"si")
			CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:10px;' ","","","","","",sLink,"si")
			'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")

			wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "<br><br></div>" & "<br>"

			wentro = 0

			strsql = "Select * From SysEstprty"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			t = 0
			Do Until rs.eof
				For ti = 1 To 18
					If (rs(ti) & "") = "" Then
						sti = ""
					Else
						sti = rs(ti)
					End If

					Pdef(t,ti - 1) = sti

				Next
				rs.movenext
				t = t + 1
			Loop

			Call ofv.cerrarconsulta(rs)

			strsql = "Select * From ObjEstPrty where detcon = " & bobj
			'response.write strsql & "<br>"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			t = 0
			Do Until rs.eof
				Select Case rs(4)
					Case "A"
					t = 0
					Case "B"
					t = 1
					Case "C"
					t = 2
					Case "D"
					t = 3
					Case "E"
					t = 4
					Case "F"
					t = 5
					Case "G"
					t = 6
				End Select

				For ti = 5 To 22
					If (rs(ti) & "") = "" Then
						sti = ""
					Else
						sti = rs(ti)
						Pdef(t,ti - 5) = sti
					End If



				Next
				rs.movenext
			Loop

			Call ofv.cerrarconsulta(rs)

			PInden = Pdef(0,7)
			If PInden <> "NO" Then PInden = "SI"

			PNumer = Pdef(0,12)
			If PNumer <> "NO" Then PNumer = "SI"

			PNumCls = Pdef(0,14)
			ReDim Preserve PNumTip(13)
			For i=0 To 13
				PNumTip(i) = Pdef(i,13)
				If PNumTip(i) = "" Then PNumTip(i) = "1"
			Next

			sant = "X"
			bseci = 0
			bsecix = 0

			For i=0 To 8
				bsec(i) = 0
				bnum(i) = 0
				bnumG(i) = 0
			Next

			inx = 0

			strsql = " SELECT * FROM CSESTDET "
			strsql = strsql & " where detcon = " & bobj & "  "
			strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
			strsql = strsql & " and not ( "
			strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
			strsql = strsql & " ) and stmodi = '' "

			If session("APLMDCLI")="BC" Then
				strsql=strsql & " AND DETTIPO < 'G'"
			End If
			strsql = strsql & " order by TSEC  "
			' response.write strsql

			Set Rsxx = Server.CreateObject("ADODB.Recordset")
			Rsxx.CursorType = 3
			Rsxx.CacheSize = 100

			Rsxx.Open strsql, cn


			If Not rsxx.eof Then
				ars = rsxx.GetRows()

			Else
				campos = rsxx.fields.count - 1
				ReDim  ars(CInt(campos),0)
				ars(39,0) = "N"
				wentro=1
			End If


			rsxx.Close
			Set rsxx = Nothing

			Do While wentro = 0



				For Inx = inx To UBound(ars,2)



					SCLAVE   = ars(0,inx)
					Stip     = ars(5,inx)

					'Secuencia
					Select Case stip
						Case "A"
						t = 7
						bseci = 0
						bsecix = 0
						bsec(1) = ars(4,inx)
						bnum(0) = bnum(0) + 1
						bnumG(0) = bnumG(0) + 1
						For i=bseci + 1 To 8
							bnum(i) = 0
							bnumG(i) = 0
						Next
						Case "B"
						t = 8
						bseci = 1
						bsecix = 1
						bsec(2) = ars(4,inx)
						bnum(1) = bnum(1) + 1
						bnumG(1) = bnumG(1) + 1
						For i=bseci + 1 To 8
							bnum(i) = 0
							bnumG(i) = 0
						Next
						Case "C"
						t = 9
						bseci = 2
						bsecix = 2
						bsec(3) = ars(4,inx)
						bnum(2) = bnum(2) + 1
						bnumG(2) = bnumG(2) + 1
						For i=bseci + 1 To 8
							bnum(i) = 0
							bnumG(i) = 0
						Next
						Case "D"
						t = 10
						bseci = 3
						bsecix = 3
						bsec(4) = ars(4,inx)
						bnum(3) = bnum(3) + 1
						bnumG(3) = bnumG(3) + 1
						For i=bseci + 1 To 8
							bnum(i) = 0
							bnumG(i) = 0
						Next
						Case "E"
						t = 11
						bseci = 4
						bsecix = 4
						bsec(5) = ars(4,inx)
						bnum(4) = bnum(4) + 1
						bnumG(4) = bnumG(4) + 1
						For i=bseci + 1 To 8
							bnum(i) = 0
							bnumG(i) = 0
						Next
						Case "F"
						t = 12
						bseci = 5
						bsecix = 5
						bsec(6) = ars(4,inx)
						bnum(5) = bnum(5) + 1
						bnumG(5) = bnumG(5) + 1
						For i=bseci + 1 To 8
							bnum(i) = 0
							bnumG(i) = 0
						Next
						Case "G"
						t = 13




						Select Case CStr(spsecn & "")
							Case "1"
							ta = 7
							bseci = 0
							Case "2"
							'ta = 8
							ta = 7
							bseci = 1
							Case "3"
							ta = 9
							bseci = 2
							Case "4"
							ta = 10
							bseci = 3
							Case "5"
							ta = 11
							bseci = 4
							Case "6"
							ta = 12
							bseci = 5
						End Select


						bnum(6) = bnum(6) + 1
						bnumG(6) = bnumG(6) + 1

						bsecix = bseci
						bsec(7) = ars(4,inx)

						Case "H"
						Select Case sant
							Case "A"
							bseci = 0
							Case "B"
							bseci = 1
							Case "C"
							bseci = 2
							Case "D"
							bseci = 3
							Case "E"
							bseci = 4
							Case "F"
							bseci = 5
						End Select
						bnum(7) = bnum(7) + 1
						bnumG(7) = bnumG(7) + 1
						bsecix = bseci + 1
						bsec(7) = ars(4,inx)
						Case "I"
						Select Case sant
							Case "A"
							bseci = 0
							Case "B"
							bseci = 1
							Case "C"
							bseci = 2
							Case "D"
							bseci = 3
							Case "E"
							bseci = 4
							Case "F"
							bseci = 5
						End Select
						bnum(8) = bnum(8) + 1
						bnumG(8) = bnumG(8) + 1
						bsecix = bseci + 1
						bsec(7) = ars(4,inx)
						Case "J"
						Select Case sant
							Case "A"
							bseci = 0
							Case "B"
							bseci = 1
							Case "C"
							bseci = 2
							Case "D"
							bseci = 3
							Case "E"
							bseci = 4
							Case "F"
							bseci = 5
						End Select
						bnum(8) = bnum(8) + 1
						bnumG(8) = bnumG(8) + 1
						bsecix = bseci + 1
						bsec(8) = ars(4,inx)

					End Select



					If stip <> "H" And stip <> "I"  And stip <> "J" Then

						For ti = 0 To 17
							Pdef(t,ti) = Pdef(t-7,ti)
						Next
						Pdef(t,18) = "SI"


						If ars(31,inx) = "NO" Then
							For ti = 11 To 28
								If (ars(ti,inx)&"") = "" Then
									sti = ""
								Else
									sti             = ars(ti,inx)
									Pdef(t,ti - 11) = sti
								End If



							Next
							Pdef(t,18) = "NO"

						End If
					End If

					BsecD = ""
					For i=1 To bsecix
						BsecD = BsecD & bsec(i) & "."
					Next





					'response.write strsql
					If Pdef(t,18) = "NO" Then
						tx = t
					Else
						tx = t - 7
					End If


					PNumCls = Pdef(tx,14)
					If PNumCls = "" Then PNumCls = "Simple"



					If stip <> "H" And stip <> "I"  And stip <> "J" Then
						PNumTip(tx) = Pdef(tx,13)
						If  PNumTip(tx) = "" Then PNumTip(tx) = "1"
					End If

					If stip = "G" Then
						If Pdef(tx,9) = "NO" Then
							MPvin = "NO"
						Else
							MPvin = "SI"
						End If
					End If

					' Anula la numeracion para este tipo o elemento
					If  PNumTip(tx) = "NO" Then PNumCls = PNumTip(tx)

					PFont  = Pdef(tx,0)
					PFsize = Pdef(tx,1)
					PFcol  = Pdef(tx,2)
					Pbgcol = Pdef(tx,3)
					PBorde = Pdef(tx,4)
					PFtipo = Pdef(tx,5)

					If marca = "CA" Then
						PFont  = "VERDANA"
						If smodi Then
							PFcol  = "#000080"
						Else
							PFcol  = "#aaaaaa"
						End If
						PFsize = "8"
						Pbgcol = ""
						PBorde = ""
						PFtipo = ""
					End If


					PSang  = Pdef(tx,6)
					txi = tx
					If tx > 6 Then txi = tx - 7
					PEspa  = Pdef(txi,8)



					Palign = Pdef(tx,16)
					If Palign = "" Then Palign = "left"
					Ptalig = Pdef(tx,17)


					txx = ta

					xmpvin = "NO"

					If stip = "G" And Pdef(txx,9) = "SI" Then
						xmpvin = "SI"
					End If

					If stip = "G" And MPvin  = "NO" Then
						xmpvin = "NO"
					End If

					xPNumS = "NO"

					If stip = "G" Then
						If Pnumer = "SI" Then
							Select Case UCase(Pnumcls)
								Case "SIMPLE"
								xPnumS = "SI"
								Case "COMPUESTA"
								xPnumS = "SI"
							End Select

						End If
					End If


					' TITULOS Y SUBTITULOS --------------------------------------------------

					If stip <> "G" And stip <> "H" And stip <> "I"  And stip <> "J" Then
						'Descripcion
						PNumS = ""
						If Pnumer = "SI" Then
							Select Case UCase(Pnumcls)
								Case "SIMPLE"
								Select Case PnumTip(tx)
									Case "1"
									BnumS = BnumG(bseci)
									Case Else
									BnumS = ConvNumer(BnumG(bseci),PnumTip(tx))
								End Select

								PnumS = BnumS & ". "
								Case "COMPUESTA"
								For i=0 To bseci
									Select Case PnumTip(i + 7)
										Case "NO"
										BnumS = ""
										Case "1"
										BnumS = BnumG(i) & "."
										Case Else
										BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
									End Select

									PnumS = PnumS & BnumS
								Next
								PnumS = PnumS & " "
							End Select

						End If



						'''''''''' arma tabla de indices '''''''''''''''''''''''''



						stexto1 = PnumS & Capitalize(ars(7,inx) & "")
						stexto1 = Replace(stexto1," ","&nbsp;")

						For nn=0 To (bseci+1)
							'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
							stexto1 = "&nbsp;" & stexto1
						Next

						wwsNIV2     = bobj & ""
						wwsNIV3     = (ars(0,inx) & "")

						sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

						'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
						iitemC=iitemC+1

						'Se saca border solid
						sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
						iitemC=iitemC+1

						sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
						Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

						'CTABELEN = ofv.GenCelda("","UT  style='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
						CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
						'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,sikon,"",CTABELEN,"si")
						wwwpad = "20"
						'if stip <> "A" then wwwpad = "30"
						wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:20px;background-color:" & sikon & ";padding-left:" & wwwpad & "px;'>" &  sLink & "<br></div>" & "<br>"

						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




					End If



					' PARRAFOS DE TEXTO    --------------------------------------------------

					If stip = "G" Then

						PNumS = ""
						If Pnumer = "SI" Then
							Select Case UCase(Pnumcls)
								Case "NO"
								BnumG(6) = BnumG(6) - 1

								Case "SIMPLE"
								Select Case PnumTip(tx)
									Case "1"
									BnumS = BnumG(6)
									Case Else
									BnumS = ConvNumer(BnumG(6),PnumTip(tx))
								End Select

								PnumS = BnumS & ". "
								Case "COMPUESTA"
								For i=0 To 6
									If BnumG(i) > 0 Then
										Select Case PnumTip(i + 7)
											Case "NO"
											BnumS = ""
											Case "1"
											BnumS = BnumG(i) & "."
											Case Else
											BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
										End Select
										PnumS = PnumS & BnumS
									End If
								Next
								PnumS = PnumS & " "
							End Select

							If UCase(Pnumcls) <> "NO" Then

								sTextoV =  ars(8,inx)

								If Len(sTextoV) > 50 Then sTextoV =  Left(sTextoV,50) & "..."

								'''''''''' arma tabla de indices '''''''''''''''''''''''''



								stexto1 = PnumS & Capitalize(sTextoV & "")
								stexto1 = Replace(stexto1," ","&nbsp;")

								For nn=0 To (bseci+1)
									'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
									stexto1 = "&nbsp;" & stexto1
								Next

								wwsNIV2     = bobj & ""
								wwsNIV3     = (ars(0,inx) & "")

								sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

								'sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
								iitemC=iitemC+1

								'Se saca border solid
								sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  &" style='border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
								iitemC=iitemC+1

								sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
								Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

								CTABELEN = ofv.GenCelda("","UT  style='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")

								'FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")
								wwwbloke = wwwbloke & "<div style='border:solid;border-width:1px;border-radius:10px;background-color:" & sikon & ";'><br>" &  sLink & "<br><br></div>" & "<br>"

								'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

							End If

						End If

					End If

					'------------------------------------------------


					Select Case stip
						Case "A"
						sant     = "A"
						sVinoAnt = sant

						Case "B"
						sant     = "B"
						sVinoAnt = sant
						Case "C"
						sant     = "C"
						sVinoAnt = sant
						Case "D"
						sant     = "D"
						sVinoAnt = sant
						Case "E"
						sant     = "E"
						sVinoAnt = sant
						Case "F"
						sant     = "F"
						sVinoAnt = sant
						Case Else

						sVinoAnt = stip

					End Select

				Next

				If inx > UBound(ars,2) Then
					Exit Do

				End If

				ssale = "no"
				Do Until ssale = "si"

					If     sant = "A" Then

						sant = "X"
					ElseIf sant = "B" Then

						sant = "A"
					ElseIf sant = "C" Then

						sant = "B"
					ElseIf sant = "D" Then

						sant = "C"
					ElseIf sant = "E" Then

						sant = "D"
					ElseIf sant = "F" Then

						sant = "E"
					End If

					If sant = "X" Then
						ssale = "si"

					End If

				Loop

			Loop

			rsxx1.movenext
		Loop

		wwwdisplay = "none"

		If (sSClick & "") = (gtsolapa & "") Then wwwdisplay = "block"

		'Se saca border solid y size y se agrega radius & sSClick & " "
		wwwbloke1 = "<div class='BLK' id='BLOK" & gtsolapa & "' style='display:" & wwwdisplay & ";' >" & wwwbloke & "</div>"

		CTABELEN = ofv.GenCelda("","UT  style='border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",wwwbloke1 ,"si")

		FTABELENX = FTABELENX & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

		Call ofv.cerrarconsulta(rsxx1)


		Case "12"

		Case "13"

		Case "14"

		Case "15"

		Case "16"



	End Select

	'' FTABELEN = FTABELEN & FTABELENX
	'' GetIndice = "ok"

	GetIndice = FTABELENX


End Function


Function array_merge(arr, arr2)

	If ISARRAY(ARR2) Then
		For Each elm In arr2
			ReDim preserve arr(UBound(arr) + 1)
			arr(UBound(arr)) = elm
		Next
	End If
	array_merge = arr
End Function


Function FindInArray(Arr,cb,db,cd,dr)

	FindInArray=dr

	For i= 0 To UBound(Arr, 2)

		If cstr(Arr(cb,i) & "")= cstr(db & "") Then
			FindInArray=Arr(cd,i)
			Exit For
		End If

	Next

End Function

Function TitMenuNovedades(P1)

	Select Case ucase(p1)

		Case "VIEW1"
		If session("APLMDCLI")="BGAL" Then
			TitMenuNovedades=""
		Else
			TitMenuNovedades="<font size=5>&nbsp;&nbsp;&nbsp;Novedades</font>"
		End If

		Case "VIEW2"
		TitMenuNovedades="<font size=5>&nbsp;&nbsp;&nbsp;Favoritos</font>"
		Case "VIEW3"
		TitMenuNovedades="<font size=5>&nbsp;&nbsp;&nbsp;Suscripciones</font>"
		Case "VIEW4"
		TitMenuNovedades="<font size=5>&nbsp;&nbsp;&nbsp;Autorizaci n</font>"
		Case Else
		If session("APLMDCLI")="BGAL" Then
			TitMenuNovedades=""
		Else
			TitMenuNovedades="<font size=5>&nbsp;&nbsp;&nbsp;Novedades</font>"
		End If
	End Select

End  Function

Function TipoSite(p1)

	' Devuelve  "HTTP:// o "HTTPs://
	TipoSite="HTTP://"

	Select Case session("APLMDCLI")

		Case "BP"
		TWWW=instr(1,ucase(Request.ServerVariables ("HTTP_REFERER")),"HTTPS")
		If TWWW > 0  Then TipoSite="HTTPS://"

		Case Else
		TWWW=ucase(Request.ServerVariables ("HTTPS"))

		Select Case  TWWW
			Case "ON"
			TipoSite="HTTPS://"

		End Select

	End Select
End Function

Function CambioInicio(P1)
	Select Case P1
		Case "BGAL"
		CambioInicio="OP1"
		Case Else
		CambioInicio="OPO"
	End Select

End Function





function DocFormat(sheader,scontenido,sfooter,hclass,cclass,fclass,tipo,xtmas)

xcontenedorf=""
select case tipo
	case "tema"
 '    xcontenedorf="<div asclass='contenedor'  >"
     'header
  '   xcontenedorf=xcontenedorf &     "<div class='" & hclass & "'>"& sheader &"</div>"
     'body
	' xcontenedorf=xcontenedorf &     "<div class='" & cclass & "'  >"

	 xcontenedorf=xcontenedorf & scontenido

	' xcontenedorf=xcontenedorf &     "</div>"
	 'footer
	'xcontenedorf=xcontenedorf &     "<div class='" &  fclass &"'>"
	'xcontenedorf=xcontenedorf &       "Última actualización: Abr 2025"
	'xcontenedorf=xcontenedorf &     "</div>"

	' xcontenedorf=xcontenedorf &"</div>"

end select
DocFormat = xcontenedorf

end function

Function SoloNumeros(strEntrada)
    Dim objRegExp, strLimpio

    ' Si el string está vacío, devolvemos vacío para evitar errores
    If IsNull(strEntrada) Or strEntrada = "" Then
        SoloNumeros = ""
        Exit Function
    End If

    ' Creamos el objeto de Expresiones Regulares
    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = "[^0-9]" ' El patrón: busca todo lo que NO sea número
        .IgnoreCase = True
        .Global = True      ' Que busque en todo el string, no solo el primero
    End With

    ' Reemplazamos lo que no es número por nada ""
    strLimpio = objRegExp.Replace(strEntrada, "")

    Set objRegExp = Nothing
    SoloNumeros = strLimpio
End Function

Function GetVersion(virtualPath)
   Dim fso, f, p, sessionName, versionedPath

    ' 1. Convertimos la ruta virtual a física
    On Error Resume Next ' Prevenir caída si la ruta está mal

    ' Creamos un nombre único para la variable de sesión basado en la ruta
	    ' Ejemplo: "cache_js_validaciones_js"
	    sessionName = "cache_" & Replace(Replace(virtualPath, "/", "_"), ".", "_")
    '    response.write sessionName
	    ' 1. Intentar obtenerlo de la sesión
	    versionedPath = Session(sessionName)

	    ' 2. Si la sesión está vacía, calculamos
	    If versionedPath = "" Then
	        p = Server.MapPath(virtualPath)
	        Set fso = Server.CreateObject("Scripting.FileSystemObject")

	        If fso.FileExists(p) Then
	            Set f = fso.GetFile(p)
	            ' Guardamos la ruta original + el nombre + la versión
	            ' Usamos f.DateLastModified para el versionado
	            versionedPath =  SoloNumeros(f.DateLastModified & "")
				'response.write versionedPath
	            ' Guardamos en sesión para la próxima vez
	            Session(sessionName) = versionedPath
	        Else
	            versionedPath = virtualPath ' Si no existe, devolvemos la ruta normal
	        End If

	        Set f = Nothing
	        Set fso = Nothing
	    End If

    		GetVersion = versionedPath


    		On Error GoTo 0
End Function

Function GetFileTimestamp(fileName)
    Dim fso, f, path
    path = Server.MapPath(fileName)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        Set f = fso.GetFile(path)
        ' Retorna la fecha convertida a número para la URL
        GetFileTimestamp = Replace(Replace(Replace(f.DateLastModified, " ", ""), ":", ""), "/", "")
    Else
        GetFileTimestamp = Timer()
    End If
    Set f = Nothing
    Set fso = Nothing
End Function


function GenEmpresas(ATREMP,p1,p2)
        DiM fila,celda,shref,tabla
		Fila    = ""


		Fila    = ""
		Celda   = ""
		sHref   = ""

		sImage  = ofv.GenImage("","x style='max-width:9rem;'","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
		Celda   = Celda & ofv.GenCelda("","t style='max-width:9rem;'","center","5%","","","",sLink,"si")

		contenido = ""' "<b>" & ATREMP(0,4) & "</b>&nbsp;"
		FOR IAEMP=1 TO TAEMP
		IF ATREMP(IAEMP,2) = 1 THEN
			'IF ATREMP(IAEMP,0) <> ATREMP(0,0) THEN
			select case p1
			case "CONSDOC"
			sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
			contenido2 = "<font style='font-size:0.7rem;'>&nbsp;BUSCANDO EN:&nbsp;<font class='noved-doc-tag'>" &ATREMP(0,4) & "</font></font>"
			CASE "CONSCMP"
			sHref = "SubconsultaComponentes.asp?EMPRE=" & ATREMP(IAEMP,0) & "&Nombre=" & sNombren & "&Menu=" & sMenu
			contenido2 = "<font style='font-size:0.7rem;'>&nbsp;BUSCANDO EN:&nbsp;<font class='noved-doc-tag'>" &ATREMP(0,4) & "</font></font>"
			CASE "OBJTREE"
				sHref = "OBJtree.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
				contenido2 = "<font style='font-size:0.7rem;'>&nbsp;ESTRUCTURA:&nbsp;<font class='noved-doc-tag'>" &ATREMP(0,4) & "</font></font>"
			CASE "NOVEDTIT"
			    sHref = "NovedTit.asp?EMPRE=" & ATREMP(IAEMP,0)
				contenido2 = "<font style='font-size:0.7rem;'>&nbsp;NOVEDADES:&nbsp;<font class='noved-doc-tag'>" &ATREMP(0,4) & "</font></font>"
			CASE "IMGFILES"
				sHref = "ImgFiles.asp?EMPRE=" & ATREMP(IAEMP,0)
				contenido2 = "<font style='font-size:0.7rem;'>REPOSITORIO&nbsp;DE&nbsp;IMAGENES:&nbsp;<font class='noved-doc-tag'>" &ATREMP(0,4) & "</font></font>"
			END SELECT
			    if ATREMP(IAEMP,0)=session("un") then
			     contenido =contenido  & "<font><a class='selec' href='" & sHref & "'><b><i>" & ATREMP(IAEMP,4) & " |</a></i></b> </font>"
			    else
				 contenido =contenido  & "<font><a class='selec' href='" & sHref & "'>" & ATREMP(IAEMP,4) & " |</a> </font>"
				end if
			'end if
		end if
		next

		Celda   = Celda & ofv.GenCelda("","","","","","","", contenido,"si")



		Celda   = Celda & ofv.GenCelda("","","right","","","","", contenido2,"si")

		Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

		GenEmpresas   = ofv.GenTabla("","aaa style='background-color:color-mix(in srgb, #132574, transparent 90%);' ","","80%","","0","","-1","-2","","",Fila,"si")

end function

%>