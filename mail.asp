<!-- #INCLUDE FILE="coremanager/IncFile/CorConect.inc" -->
<%
 'smtp.gmail.com


  stowho	 ="soptec@corconsulting.com.ar"
  FROM		 ="soptec@corconsulting.com.ar"

'  Set objCDOSYSMailq = Server.CreateObject("System.Net.Mail")
  message=" prueba de envio de mail " & date() & "  "  & time()

  Xmessage = GenMDLMAIL (stowho,FROM,message)
  subject="Prueba"
  CALL SendMailEX(Xmessage, stowho, FROM,subject, session("MailComponent"),"SI",0, "NO")
  response.write  "MailComponent=" & session("MailComponent") & "<br>"
  response.write  "SMTPServer=" & Session("SMTPServer") & "<br>"
  response.write  "smtpport=" & Session("SMTPport") & "<br>"
  response.write  "Err.Number=" & Err.Number & "<br>"
  response.write  "Err.Description=" & Err.Description & "<br>"

Function SendMailEX(ByVal strEmailBodyMessage, ByVal strRecipientEmailAddress , ByVal strFromEmailAddress, ByVal strSubject, strMailComponent, blnHTML,ByVal Importance, ByVal notification)

	'Dimension variables
	Dim objCDOSYSMail		'Holds the CDOSYS mail object
	Dim objCDOMail			'Holds the CDONTS mail object
	Dim objJMail			'Holds the Jmail object
	Dim objAspEmail			'Holds the Persits AspEmail email object
	Dim objAspMail			'Holds the Server Objects AspMail email object
	Dim strEmailBodyAppendMessage	'Holds the appended email message
	Dim SmtpServer
	Dim SmtpPort
    On Error Resume Next
	SmtpServer = Session("SMTPServer")
	SmtpPort   = Session("SMTPPort")

	' RESPONSE.WRITE strRecipientEmailAddress & "<BR>" & strFromEmailAddress
	SHTMLMail = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
	SHTMLMail = "<html>"
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
		Case "CDOSYSww"





		schema = "http://schemas.microsoft.com/cdo/configuration/"

		Set msg = CreateObject("CDO.Message")
		msg.Subject  = "Test"
		msg.From     = "soptec@gmail.com"
		msg.To       = "soptec@gmail.com"
		msg.TextBody = "This is some sample message text."

		With msg.Configuration.Fields
		  .Item(schema & "sendusing")      = 2
		  .Item(schema & "smtpserver")     = "smtp.gmail.com"
		  .Item(schema & "smtpserverport") = 465
		  .Item(schema & "smtpauthenticate") = 1
		  .Item(schema & "sendusername") = "soptec@corconsulting.com.ar"

		  .Item(schema &"sendpassword") = "Pichon1612"

		  .Update
		End With

		msg.Send


Case "CDOSYS"
		'Dimension variables
		Dim objCDOSYSCon

		'Create the e-mail server object
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

		'Set and update fields properties
'inetmgr6.exe
		'Out going SMTP server
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
		'SMTP port
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  =25
		'CDO Port
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

        ' esto es por google no va en los clientes ---------------------------------------
		'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true


	'	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "soptec@corconsulting.com.ar"
	''	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpaccountname")   = "soptec@corconsulting.com.ar"
	'   objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Pichon1612"
		'objCDOSYSCon-------------------------------------------------------------------------------------






		'Timeout
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		objCDOSYSCon.Fields.Update


		'Update the CDOSYS Configuration
		Set objCDOSYSMail.Configuration = objCDOSYSCon


		'Who the e-mail is from
		'objCDOSYSMail.From = strFromEmailName & " <" & strFromEmailAddress & ">"
		response.write strFromEmailAddress & "<br>"
		objCDOSYSMail.From = strFromEmailAddress

		'Who the e-mail is sent to
		'objCDOSYSMail.To = strRecipientName & " <" & strRecipientEmailAddress & ">"
		objCDOSYSMail.To =  strRecipientEmailAddress
		response.write strFromEmailAddress & "<br>"

		'The subject of the e-mail
		objCDOSYSMail.Subject = strSubject

		'Set the e-mail body format (HTMLBody=HTML TextBody=Plain)


		'objCDOSYSMail.AutoGenerateTextBody = false


		If UCase(blnHTML)="SI" Then
		'objCDOSYSMail.TextBody = SHTMLMail
			objCDOSYSMail.HTMLBody = SHTMLMail
		Else
			objCDOSYSMail.TextBody = SHTMLMail
		End If



		objCDOSYSMail.Fields.Update

		'Send the e-mail

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
	SendMailEX = True
End Function
 response.write " Se envio mail"
%>