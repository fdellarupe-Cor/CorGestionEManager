<%

' ****************************************************************************************
' ** Cor E-Solution v3 -    Esquema de Configuracion ------------------ marzo, 2007     **
' ****************************************************************************************


''XLCHARSET = "iso-8859-1"
  Const ORDERTYPE_ASC   = 10
  Const ORDERTYPE_DESC  = 11
  Const SORTBY_KEY      = 20
  Const SORTBY_VALUE    = 21

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
DIM ATREMP(50,5)

USREMP = 1

DIM ATXCOM(20,4)
GLBCOMBO    = 0
GLBCACTU    = 0


dim ATABGLOSLNKS(500,2)
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


  if Session("CorApli") = "COREM20" then
     wkaplic = "EM"
  elseif Session("CorApli") = "COREP20" then
     wkaplic = "PL"
  elseif Session("CorApli") = "CLI916953113116518847748" then
     wkaplic = "GP"
  elseif Session("CorApli") = "CORESS20" then
     wkaplic = "CE"
  elseif Session("CorApli") = "CLI111557214260398409701" then
     wkaplic = "OR"
  else
     wkaplic = "RG"
  end if



 'SmtpServer = "smtp.iplannetworks.net"
 'SmtpPort   = 25


if session("Usuario") = "" then
      session("LUID") = ofv.ADbUid
      session("LPWD") = ofv.ADbPwd
end if

ofv.ADbUidl = session("LUID")
ofv.ADbPwdl = session("LPWD")


  if request.querystring("MNN") <> "" then
     session("FormN") = replace(request.querystring("MNN"),"_"," ")
  end if

  if request.querystring("seg") <> "" then
     session("Form") = replace(request.querystring("seg"),"_"," ")
  end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Funciones de administracion de Windows
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function GTkey(XCodobJ)

GThms   = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses   = session.sessionid
GTsecu  = cstr(XCodobJ & GThms & GTses)

GTsecuf = ""

for GTi=1 to len(GTsecu)
    GTcs = mid(GTsecu,GTi,1)
    if  GTcs => "0" and GTcs <= "9" then
        GTsecuf = GTsecuf & GTcs
    end if
next

GTkey  = trim(GTsecuf)

end function

Function ConsCambios(sASP)

     Dim wxs
     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " function AbrirCambios()" & Chr(13)
     wxs = wxs & " {" & Chr(13)
      if BrowserName() = "IE" Then
       wxs = wxs & " window.showModalDialog(" & Chr(34) & sASP & Chr(34) & ",window,"
       wxs = wxs &  Chr(34) & "dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"
     wxs = wxs & "dialogTop:250px;dialogLeft:300px;status:no" & Chr(34)
     else
      wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34) & ",window,"
      wxs = wxs &  Chr(34) & "status=no,Width=600,Height=350,top=250,left=300"
     wxs = wxs & Chr(34)
     end if


     wxs = wxs & " ); " & Chr(13)
     wxs = wxs & "}" & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

 ConsCambios =      wxs

End Function

Function Ordenar(sASP)



     Dim wxs

     wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bor1", "ORdenar", "button", "", "10", "bt", "AbrirOrdenar()") & "<br><br>", "si")
     wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "function AbrirOrdenar() { " & Chr(13)
     wxs = wxs & " var h=570;w=790 ;"
     if  session("APLMDCLI")="BC" THEN
     wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
     END IF
     if BrowserName()="IE" then
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
     else
     wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
    end if'
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
     wxs = wxs & "function AbrirVentana() { " & Chr(13)
     wxs = wxs & " var h=570;w=790 ;"
	   if  session("APLMDCLI")="BC" THEN
	    wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
		END IF
    if BrowserName()="IE" then
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
     else
     wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
    end if
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
     wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
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
        if  session("APLMDCLI")="BC" THEN
         wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
          END IF
    if BrowserName()="IE" then
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
     else
     wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
    end if'
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
     wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
     wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

     Eliminar =      wxs


End Function

Function Copiar(sASP)



     Dim      wxs

     wxs = ofv.GenCelda("", "", "center", "33%", "", "", "", ofv.GenerarInput("bco1", "Copiar", "button", "", "10", "bt", "AbrirCopiar()") & "<br><br>", "si")
     wxs = ofv.GenForm("", "", "", "", "post", "",wxs, "si")

     wxs = wxs & "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "function AbrirCopiar() {" & Chr(13)
     wxs = wxs & " var h=570;w=790 ;"
        if  session("APLMDCLI")="BC" THEN
         wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
          END IF
    if BrowserName()="IE" then
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
     else
     wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
    end if'
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
     wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
     wxs = wxs & " ); " & Chr(13) & "window.close(); " & Chr(13)
     wxs = wxs & " } " & Chr(13) & "</script>" & Chr(13)

     Copiar =      wxs


End Function

FUNCTION MOVER(sASP)

   DIM wxs

    wxs = ofv.GenCelda("", "", "center", "33%", "", "", "",   ofv.GenerarInput("bco1", "Mover", "button", "", "10", "bt", "AbrirMover()") & "<br><br>"  , "si")
    wxs = ofv.GenForm("", "", "", "", "post", "", wxs, "si")

   wxs = wxs & "<script language=JavaScript>" & CHR(13)
   wxs = wxs & "function AbrirMover() {" & CHR(13)
    wxs = wxs & " var h=570;w=790 ;"
        if  session("APLMDCLI")="BC" THEN
         wxs = wxs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
          END IF
    if BrowserName()="IE" then
     wxs = wxs & " window.dialogArguments.open(" & Chr(34) & sASP & Chr(34)
     'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
     else
     wxs = wxs & " window.opener.open(" & Chr(34) & sASP & Chr(34)
    end if'
     wxs = wxs & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
     wxs = wxs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
   wxs = wxs & " ); " & CHR(13) & "window.close(); " & CHR(13)
   wxs = wxs & " } " & CHR(13) & "</script>" & CHR(13)

   MOVER = wxs

END FUNCTION

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Funciones de administracion de EMail
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function SendMail(ByVal strEmailBodyMessage, ByVal strRecipientEmailAddress , ByVal strFromEmailAddress, ByVal strSubject, strMailComponent, blnHTML,Byval Importance, byval notification)

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
	Select Case ucase(strMailComponent)


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

 ' esto es por google no va en los clientes ---------------------------------------
					                     objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
					                    objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

								       objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "soptec@corconsulting.com.ar"
									   objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Pomelo16"
					'-------------------------------------------------------------------------------------


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

				'The subject of the e-mail
				objCDOSYSMail.Subject = strSubject

				'Set the e-mail body format (HTMLBody=HTML TextBody=Plain)
				If UCASE(blnHTML)="SI" Then
				 	objCDOSYSMail.HTMLBody = SHTMLMail
				Else
					objCDOSYSMail.TextBody = SHTMLMail
				End If
                if UCASE(notification)="SI" then
                   objCDOSYSMail.fields("urn:schemas:mailheader:disposition-notification-to") =chr34 & strFromEmailAddress & chr(34) ' "<valid e-mail address>" 'ToDo: Type a valid e-mail address.
		   objCDOSYSMail.fields("urn:schemas:mailheader:return-receipt-to") = chr34 & strFromEmailAddress & chr(34) 			' "<valid e-mail address>"  'ToDo: Type a valid e-mail address.
                end if
                select case importance
                case 0 'BAJA
                 objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 5
                case 1 'NORMAL
                  objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 3
                case 2  'ALTA
                   objCDOSYSMail.Fields("urn:schemas:mailheader:X-Priority").value= 1

                end select

				objCDOSYSMail.Fields.Update

				'Send the e-mail
				on error resume next
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

                               if UCASE(notification) = "SI" then
                                  strReply_To = "<" & From & ">"
                                  objCDOMail.Value("Disposition-Notification-To") = strReply_To
                               end if

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
				If NOT strIncomingMailServer = "" Then .Execute
			End With

			'Close the server mail object
			Set objJMail = Nothing




		'******************************************
		'*** Persits AspEmail mail component   ****
		'******************************************

		'AspEmail component
		Case ucase("AspEmail")
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
				If ucase(blnHTML) = "SI" Then
					objAspEmail.IsHTML = True
				End If

				'The main body of the e-mail
				objAspEmail.Body = SHTMLMail

				 select case importance
				   case 0 'BAJA
				    objAspEmail.Priority=5
				   case 1 'NORMAL
				    objAspEmail.Priority=3
				   case 2  'ALTA
				    objAspEmail.Priority=1
                 end select


                 if UCASE(notification)="SI" then
                  objAspEmail.AddCustomHeader "Return-Receipt-To: " & chr(34) & strFromEmailAddress & chr(34)
 				 end if


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
			   	If NOT strIncomingMailServer = "" Then .SendMail
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

function GenKeyLnks(gtlnks)

    if session("usuario")="GOOGLE" then
     GenKeyLnks="(XLNK:XLNK)"
    else
     GenKeyLnks = trim(gtlnks)
    end if

end function
function GenKeyLnks2(gtlnks)

    TTLNKS = TTLNKS + 1

    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

    XCodobJ = sh & sm & ss

    GThms   = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
    GTses   = session.sessionid
    GTsecu  = cstr(XCodobJ & GThms & GTses)

    GTsecuf = "(XLNK:" & TTLNKS & ":"


    for GTi=1 to len(GTsecu)
        GTcs = mid(GTsecu,GTi,1)
        if GTcs => "0" and GTcs <= "9" then
           GTsecuf = GTsecuf & GTcs
        end if
    next


    GTsecuf = GTsecuf & "XLNK)"

    ATABGLOSLNKS(TTLNKS,1) = trim(GTsecuf)
    ATABGLOSLNKS(TTLNKS,2) = gtlnks

    GenKeyLnks2 = trim(GTsecuf)

end function



function Capitalize (str)
  Capitalize=UCase(Left(str, 1)) & LCase(Mid(str, 2))
end function



function VerGlosario(XEhtml)

  VerGlosario = XEhtml

end function

function armaglos()
dim  RSX1zz
cadena=""
strsqlP = "Select KEYWORD,SHWWORD From GLOSARIO ORDER BY SHWWORD"
 set cccn = ofv.conectar(ofv.strconn2)

  set rsX1zz = ofv.crearconsultaEx(StrSqlP,cccn,1,parametros)
  i=0

     DO WHILE NOT RSX1zz.EOF

     cadena= cadena & "Aglos[" & i &"]='" & rsx1zz(0) &"';" & chr(13)
     rsX1zz.movenext
     i=i+1
     loop
 armaglos=cadena

	  if session("yider")<> "" then
  cadena=""
  session("yider")=replace(session("yider"),"*" ,"")
  ayider=split(session("yider")," ")
  if isarray(ayider) then
  zdelim= ""
  pals=""
   'i=0
   for ia=0 to  -1 'ubound(ayider)
      pal= trim(ayider(ia) & "")
      if len(pal) > 3 then
      cadena= cadena & "Aglos[" & i &"]='" & pal &"';" & chr(13)
      i=I+1
      end if
    ' response.write pal
   next
    armaglos=armaglos & cadena
  end if

 CALL OFV.CERRARCONSULTA (rsX1zz)
 CALL OFV.CERRARCONN(cccn)
 end if
' response.write session("yider")
end function

function VerGlosario2(XEhtml)

 dim ATABGLOS(200,5)
 ttglos = 0

 XEconvMl = "<span CLASS=gls onclick=" & chr(34) & "verglos2('XXARG')" & chr(34) & " >XXDATA</span>"
' XEconvMl = "XXDATA"

' glosario

  strsql = "Select KEYWORD,SHWWORD From GLOSARIO ORDER BY SHWWORD"
  set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsX1.eof then
     DO WHILE NOT RSX1.EOF

        TTGLOS = TTGLOS + 1
        ATABGLOS(TTGLOS,0) = RSX1(0)

        ATABGLOS(TTGLOS,1) = RSX1(1)

        if  UCASE(ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) then
           ATABGLOS(TTGLOS,2) = "XXXXXX"
        else
           ATABGLOS(TTGLOS,2) = UCASE(ATABGLOS(TTGLOS,1))
        end if

        if  LCASE(ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) then
           ATABGLOS(TTGLOS,3) = "XXXXXX"
        else
           ATABGLOS(TTGLOS,3) = LCASE(ATABGLOS(TTGLOS,1))
        end if

        if  Capitalize (ATABGLOS(TTGLOS,1)) = ATABGLOS(TTGLOS,1) then
           ATABGLOS(TTGLOS,4) = "XXXXXX"
        else
           ATABGLOS(TTGLOS,4) = Capitalize (ATABGLOS(TTGLOS,1))
        end if





     RSX1.MOVENEXT
     LOOP

  end if
  call ofv.cerrarconsulta(rsX1)

  IF TTGLOS > 0 THEN

     FOR IGLOS = 1 TO TTGLOS
         XEconvBs = replace(XEconvML,"XXARG","XXARG" & (999-IGLOS))
         FOR III = 1 TO 4
             if  ATABGLOS(IGLOS,III) <> "XXXXXX" then
                 XEconv = replace(XEconvBs,"XXDATA",ATABGLOS(IGLOS,III))

                 XEhtml = REPLACE(XEhtml, ATABGLOS(IGLOS,III),XEconv)
             end if
         NEXT
     NEXT

     FOR IGLOS = 1 TO TTGLOS
         XEhtml   = REPLACE(XEhtml,"XXARG" & (999-IGLOS),ATABGLOS(IGLOS,0))
     NEXT

  END IF

  for IGLOS = 1 TO TTLNKS
      XEhtml   = REPLACE(XEhtml, ATABGLOSLNKS(IGLOS,1),ATABGLOSLNKS(IGLOS,2))
  next

  xehtmln = instr(XEhtml,"<span CLASS=gls")

  xehtmlx = ""


  if xehtmln = 0 then

     XEhtml = replace(XEhtml,"versup() { ","versup() { alert('No se encontraron entradas del GLOSARIO en la pagina actual'); " )

  end if


  VerGlosario2 = XEhtml

end function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCION CONFECCION DEL EMAIL ( HTML O TEXTO)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 FUNCTION GenMDLMAIL (XGUPARA,XGUORIGEN,XGUMENSAGE)

''message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"


     XGUMENSAGEw = XGUMENSAGE & " "

     XGUMENSAGEw = replace(XGUMENSAGEw,"@@@NORMREQ@@@","<u>" & SESSION("NOMREQ") & "</u>")
     XGUMENSAGEw = replace(XGUMENSAGEw,"@@@FECREQ@@@",now())



     Whtml = ""

     SELECT CASE  SESSION("MAILOPC")

            CASE "HTML"

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
                  Whtml = Whtml & "COR&nbsp;E-SOLUTION&nbsp;SUITE&nbsp;V4.6</i></B></FONT></P>"
                  Whtml = Whtml & "</DIV>"



            CASE ELSE


                  Whtml = Whtml & SESSION("NOMAPLI")  & "<BR>" & CHR(13) & CHR(13)

                  Whtml = Whtml & XGUPARA & ": <BR>" & CHR(13)

                  Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & CHR(13)

                  Whtml = Whtml & XGUMENSAGEw  & "<BR>" & CHR(13)
                  Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & CHR(13)




     END SELECT


     GenMDLMAIL = Whtml



 END FUNCTION

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCIONES DE BUSQUEDAS MULTIPLES ( USUARIOS-GRUPOS-SECTORES-EMPRESAS)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 function  GenBuscaUsrex(XGUM1,XGUM2,XGUM3,XGUM4,XGUM5,XGUM6,XGUM7,XGUM8,XGUM9,XGUM10,XGUM1F,XGUM2F,XGUM3F,XGUM4F,XGUM5F,XGUM6F,XGUM7F,XGUM8F,XGUM9F,XGUM10F)

  XGUM1FL = REPLACE(XGUM1F,CHR(9),"")
  XGUM2FL = REPLACE(XGUM2F,CHR(9),"")
  XGUM3FL = REPLACE(XGUM3F,CHR(9),"")
  XGUM4FL = REPLACE(XGUM4F,CHR(9),"")

  XGUM5FL = REPLACE(XGUM5F,CHR(9),"")
  XGUM6FL = REPLACE(XGUM6F,CHR(9),"")
  XGUM7FL = REPLACE(XGUM7F,CHR(9),"")
  XGUM8FL = REPLACE(XGUM8F,CHR(9),"")

  XGUM9FL = REPLACE(XGUM9F,CHR(9),"")
  XGUM10FL = REPLACE(XGUM10F,CHR(9),"")



  XGUFBURL = request.servervariables("SCRIPT_NAME")

  XGUFBURL = RIGHT(XGUFBURL,LEN(XGUFBURL) - (INSTRREV(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

  XGUMASTER = REQUEST.FORM("XGUMASTER")
  XGUFILTRO = REQUEST.FORM("XGUFILTRO")
  XGUITEMCD = REQUEST.FORM("XGUITEMCD")
  XGUCODIGO = REQUEST.FORM("XGUCODIGO")
  XGUCODDSC = REQUEST.FORM("XGUCODDSC")
 'RESPONSE.WRITE XGUCODDSC & "22222222"
if XGUITEMCD="" then XGUITEMCD="NO"
IF XGUCODIGO="" THEN XGUCODIGO="NO"
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


  IF SESSION("EMPMUN") = "S" then

  IF XGUM1 = "S" THEN

     IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
         XGUMASTERW = "E"
     END IF

     XGUMSBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESAS'  CLASS=BT ONCLICK=""clactfrm('E','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
     XGUMSBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI1 & "</FORM>"

  END IF

  END IF

  IF XGUM2 = "S" THEN

     IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
         XGUMASTERW = "S"
     END IF

     XGUMSBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTORES'  CLASS=BT ONCLICK=""clactfrm('S','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
     XGUMSBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI2 & "</FORM>"


  END IF

  IF XGUM3 = "S" THEN

     IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
         XGUMASTERW = "G"
     END IF

     XGUMSBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPOS'  CLASS=BT ONCLICK=""clactfrm('G','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
     XGUMSBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI3 & "</FORM>"


  END IF

  IF XGUM4 = "S" THEN

     IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
         XGUMASTERW = "U"
     END IF

     XGUMSBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('U','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
     XGUMSBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI4 & "</FORM>"


  END IF

  IF XGUM5 = "S" THEN

       IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
           XGUMASTERW = "DEPENDENCIA"
       END IF

       XGUMSBOTi5 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('DEPENDENCIA','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
       XGUMSBOT5  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI5 & "</FORM>"


  END IF


  IF XGUM6 = "S" THEN

         IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
             XGUMASTERW = "PORPUESTO"
         END IF

         XGUMSBOTi6 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORPUESTO','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
         XGUMSBOT6  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI6 & "</FORM>"


  END IF


   IF XGUM7 = "S" THEN

           IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
               XGUMASTERW = "PORUSR"
           END IF

           XGUMSBOTi6 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORUSR','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
           XGUMSBOT6  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI6 & "</FORM>"


  END IF

   IF XGUM8 = "S" THEN

             IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
                 XGUMASTERW = "PORSECPERF"
             END IF

             XGUMSBOTi7 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORSECPERF','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
             XGUMSBOT7  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI7 & "</FORM>"


  END IF


   IF XGUM9 = "S" THEN

               IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
                   XGUMASTERW = "PORSECTOREX"
               END IF

               XGUMSBOTi8 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LIMPIAR'  CLASS=BT ONCLICK=""clactfrm('PORSECTOREX','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
               XGUMSBOT8  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI8 & "</FORM>"


  END IF
    IF XGUM10 = "S" THEN

                  IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
                      XGUMASTERW = "PORGRUPOEXT"
                  END IF

                  XGUMSBOTi9 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPOS DE CORREO'  CLASS=BT ONCLICK=""clactfrm('PORGRUPOEXT','NO','NO','NO','NO');document.XGUFRM.submit();return false;"" ></td>"
                  XGUMSBOT9  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI9 & "</FORM>"


   END IF






  IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
         XGUMASTER = XGUMASTERW
  END IF


  XGUCONTEMS = ""

  XGUCONTEMS = XGUCONTEMS & "<table id=""MTBSUSR0B""  class='aaa'  cellspacing=1 cellPADDING=0 > "
  XGUCONTEMS = XGUCONTEMS & "<tr>"
  XGUCONTEMS = XGUCONTEMS & "<td class=ut WIDTH=60%  >&nbsp;</td>"

  IF XGUMSBOT1 <> "&nbsp;" THEN
     XGUCONTEMS = XGUCONTEMS & XGUMSBOT1
  END IF

  IF XGUMSBOT2 <> "&nbsp;" THEN
     XGUCONTEMS = XGUCONTEMS & XGUMSBOT2
  END IF

  IF XGUMSBOT3 <> "&nbsp;" THEN
     XGUCONTEMS = XGUCONTEMS & XGUMSBOT3
  END IF

  IF XGUMSBOT4 <> "&nbsp;" THEN
     XGUCONTEMS = XGUCONTEMS & XGUMSBOT4
  END IF


   IF XGUMSBOT5 <> "&nbsp;" THEN
       XGUCONTEMS = XGUCONTEMS & XGUMSBOT5
   END IF

   IF XGUMSBOT6 <> "&nbsp;" THEN
          XGUCONTEMS = XGUCONTEMS & XGUMSBOT6
   END IF

   IF XGUMSBOT7 <> "&nbsp;" THEN
             XGUCONTEMS = XGUCONTEMS & XGUMSBOT7
   END IF

    IF XGUMSBOT8 <> "&nbsp;" THEN
                XGUCONTEMS = XGUCONTEMS & XGUMSBOT8
   END IF
   IF XGUMSBOT9 <> "&nbsp;" THEN
                   XGUCONTEMS = XGUCONTEMS & XGUMSBOT9
    END IF


  XGUCONTEMS = XGUCONTEMS & "</tr>"
  XGUCONTEMS = XGUCONTEMS & "</table>"


  XGUFNC1 = ""
  XGUFNC2 = ""


  set xgucn = ofv.conectar(ofv.strconn2)

  XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#F7AA22;cursor:pointer;' "
  XGUSTYLN = " STYLE='cursor:pointer;' "

  XGUFBUSR = XGUCODIGO

  XGUFBTAB1  = ""
  XGUFBTAB2  = ""
  XGUFBTD1  = ""
  XGUFBTD2  = ""

  IF XGUMASTER = "" OR XGUMASTER = "NO" THEN XGUMASTER = "U"

  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '' SELECCION DE BUSQUEDA MASTER - E:EMPRESAS / S:SECTORES / G:GRUPOS / U:USUARIOS              ''
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' RESPONSE.WRITE XGUMASTER
 SESSION ("XGUMASTER" )=XGUMASTER
  SELECT CASE XGUMASTER

  CASE "PORUSR"


              XGUFILDSC = "&nbsp;POR&nbsp;NOMBRE"

      XGUFBBOTi1 =""' "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
      XGUFBBOTi2 =""' "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
      XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
      XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

      XGUFBBOT1  =""' "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
      XGUFBBOT2  =""' "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
      XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
      XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"


     IF XGUFILTRO="" THEN XGUFILTRO=request.querystring("XGUFILTRO")
     IF XGUITEMCD="" THEN XGUITEMCD=request.querystring("XGUITEMCD")


   SELECT CASE XGUFILTRO




           CASE "L"

                 XGUFILDSC1 = ":&nbsp;LETRA"

		              xTrColor = "#F9F9F9"


		              IF XGUCODIGO <> "NO" THEN
		                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
		              ELSE
		                 XGUFNC2 = ""
		              END IF


		              dim ATABXGULTRz(28)

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


		              FOR XGUINX=1 TO UBOUND(ATABXGULTRz) STEP 4

		                  XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
		                  XGUFBTD1 = XGUFBTD1 & " >"

		                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

		                  IF ATABXGULTRz(XGUINX) = XGUITEMCD THEN
		                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
		                  ELSE
		                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
		                  END IF

		                  XGUFBTD1 = XGUFBTD1 & " >"

		                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
		                  XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
		                  XGUFBTD1 = XGUFBTD1 & "</TD>"


		                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

		                  IF ATABXGULTRz(XGUINX+1) = XGUITEMCD THEN
		                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
		                  ELSE
		                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
		                  END IF

		                  XGUFBTD1 = XGUFBTD1 & " >"
		                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
		                  XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
		                  XGUFBTD1 = XGUFBTD1 & "</TD>"

		                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

		                  if  ATABXGULTRz(XGUINX+2) = "NO" then

		                     XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

		                  else
		                     IF ATABXGULTRz(XGUINX+2) = XGUITEMCD THEN
		                        XGUFBTD1 = XGUFBTD1 & XGUSTYLS
		                     ELSE
		                        XGUFBTD1 = XGUFBTD1 & XGUSTYLN
		                     END IF
		                     XGUFBTD1 = XGUFBTD1 & " >"
		                     XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
		                     XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

		                  end if

		                  XGUFBTD1 = XGUFBTD1 & "</TD>"

		                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

		                  if  ATABXGULTRz(XGUINX+3) = "NO" then

		                     XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

		                  else
		                     IF ATABXGULTRz(XGUINX+3) = XGUITEMCD THEN
		                        XGUFBTD1 = XGUFBTD1 & XGUSTYLS
		                     ELSE
		                        XGUFBTD1 = XGUFBTD1 & XGUSTYLN
		                     END IF
		                     XGUFBTD1 = XGUFBTD1 & " >"
		                     XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTRz(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
		                     XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTRz(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

		                  end if

		                  XGUFBTD1 = XGUFBTD1 & "</TD>"



		                  XGUFBTD1 = XGUFBTD1 & "</TR>"


		                  if xTrColor = "#FFFFFF" then
		                     xTrColor = "#F9F9F9"
		                  ELSE
		                     xTrColor = "#FFFFFF"
		                  END IF

		              NEXT




		              IF XGUITEMCD <> "NO" THEN

		                 CALL getUsrInTEXT2(XGUITEMCD & "*")

		                 XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
		                 XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

		                 xTrColor = "#FFFFFF"

		                 FOR XGUINX=0 TO UBOUND(XGUTABKEY)


		                  XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
		                 XGUFBTD2 = XGUFBTD2 & " >"

		                 XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

		                 IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
		                    XGUFBTD2 = XGUFBTD2 & XGUSTYLS
		                 ELSE
		                    XGUFBTD2 = XGUFBTD2 & XGUSTYLN
		                 END IF

		                 XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
		                 XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
		                 XGUFBTD2 = XGUFBTD2 & "</TR>"


		                          if xTrColor = "#FFFFFF" then
		                             xTrColor = "#F9F9F9"
		                          ELSE
		                             xTrColor = "#FFFFFF"
		                          END IF


		                 NEXT




              END IF




           CASE "T"
               ' XGUITEMCD=replace(XGUITEMCD &"","'","%27")
                XGUFILDSC1 = ":&nbsp;TEXTO"



                xTrColor = "#FFFFFF"


                IF XGUCODIGO <> "NO" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                    XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD1 = XGUFBTD1 & " >"

                    XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

                    IF XGUITEMCD <> "NO" THEN
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                    ELSE
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                    END IF

                    XGUFBTD1 = XGUFBTD1 & " >"
                    XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE=""" & XGUITEMCD& """ CLASS=DT size=40 >&nbsp;&nbsp;"

                    XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
                    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD1 = XGUFBTD1 & "</TD>"

                    XGUFBTD1 = XGUFBTD1 & "</TR>"


                IF XGUITEMCD <> "NO" THEN

                   CALL getUsrInTEXT2("*" & XGUITEMCD & "*")

                   XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                   XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                   xTrColor = "#FFFFFF"

                   FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                   XGUFBTD2 = XGUFBTD2 & " >"

                   XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                   IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                      XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                   ELSE
                      XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                   END IF

                   XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & replace(XGUITEMCD &"","'","\'")  & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                   XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                   XGUFBTD2 = XGUFBTD2 & "</TR>"


                            if xTrColor = "#FFFFFF" then
                               xTrColor = "#F9F9F9"
                            ELSE
                               xTrColor = "#FFFFFF"
                            END IF


                   NEXT




              END IF

      CASE ELSE

                 XGUFILDSC1 = "&nbsp;"

                 XGUFILTRO = "NO"
                 XGUITEMCD = "NO"
                 XGUCODIGO = "NO"
              XGUCODDSC = "NO"

  END SELECT

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

  CASE "DEPENDENCIA"

                XGUFILDSC = "&nbsp;POR&nbsp;DEPENDENCIA"

                xTrColor = "#FFFFFF"
                  'RESPONSE.WRITE XGUITEMCD & "22"
                IF XGUITEMCD <> "NO" THEN
                    XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
                ELSE
                    XGUFNC1 = ""
                END IF


                IF XGUCODIGO <> "NO" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                strsql = "Select e.* From Eco e "
                strsql =strsql & "where e.eco in(select distinct l.sector from listadepuestos l where l.grafica='S' )"
                strsql =strsql & " order by e.descri"
              ''RESPONSE.WRITE strsql & " strsql "
                set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

                do until xgurs.eof


                   XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                   XGUFBTD1 = XGUFBTD1 & " >"

                   XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

                   IF CSTR(xgurs(0)) = XGUITEMCD THEN
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                   ELSE
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                   END IF

                   XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                   XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                   XGUFBTD1 = XGUFBTD1 & "</TR>"


                            if xTrColor = "#FFFFFF" then
                               xTrColor = "#F9F9F9"
                            ELSE
                               xTrColor = "#FFFFFF"
                            END IF


                xgurs.movenext
                loop

                call ofv.cerrarconsulta(xgurs)

                session("SGU_UIGuserid")=""
                session("SGU_UIGdescri")=""
                IF XGUITEMCD <> "NO" THEN

                  ' CALL getUsrInSECT(XGUITEMCD,"")
                  if XGUITEMCD=0 then XGUITEMCD=-1
                  SQL="SELECT * FROM LISTADEPUESTOS WHERE grafica='S' and sector=" &  XGUITEMCD
                 'response.write sql
                  set xpuestos = ofv.crearconsultaEx(Sql,xgucn,1,volver)
                   delim=""

                   do while not xpuestos.eof

                   session("SGU_UIGuserid")= session("SGU_UIGuserid")   & delim  & xpuestos(0)
                   session("SGU_UIGdescri")= session("SGU_UIGdescri")  & delim  & xpuestos(1)
                   delim =chr(9)
                   xpuestos.movenext
                   loop
                   call  ofv.cerrarconsulta(xpuestos)
                   XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                   XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                   xTrColor = "#FFFFFF"

                   FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                   XGUFBTD2 = XGUFBTD2 & " >"

                   XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR PUESTO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                   IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                      XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                   ELSE
                      XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                   END IF

                   XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                   XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                   XGUFBTD2 = XGUFBTD2 & "</TR>"


                            if xTrColor = "#FFFFFF" then
                               xTrColor = "#F9F9F9"
                            ELSE
                               xTrColor = "#FFFFFF"
                            END IF


                   NEXT




              END IF
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


       CASE "PORSECTOREX"

  XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
  XGUFBBOT2  = ""
  XGUFBBOT3  = ""
  XGUFBBOT4  = ""

              XGUFILDSC = "&nbsp;SECTOR"

              xTrColor = "#FFFFFF"

              IF XGUCODIGO <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF



              STRSQL= "Select A.OBJID,B.DESCRI From objsector A"
			  STRSQL = STRSQL & " INNER JOIN ECO B ON(B.ECO= A.CODSECTOR)"
			  STRSQL=STRSQL & " WHERE A.CODOBJETO=0" & session("NCliente")
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR SECTOR.'  ID='***" & xgurs(0) & "' "

                 IF (XGUM2FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF



                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM2FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)


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
      CASE "PORGRUPOEXT"




	              XGUFILDSC = "&nbsp;POR&nbsp;GRUPOS DE CORREO"

	      XGUFBBOTi1 =""' "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	      XGUFBBOTi2 =""' "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	     ' XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
	      XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

 	      XGUFBBOT1  =""' "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
 	      XGUFBBOT2  =""' "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
 	     ' XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
 	      XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"


 	     IF XGUFILTRO="" THEN XGUFILTRO=request.querystring("XGUFILTRO")
 	     IF XGUITEMCD="" THEN XGUITEMCD=request.querystring("XGUITEMCD")


 	   SELECT CASE XGUFILTRO









 	           CASE "T"
 	               ' XGUITEMCD=replace(XGUITEMCD &"","'","%27")
 	                XGUFILDSC1 = ":&nbsp;TEXTO"



 	                xTrColor = "#FFFFFF"


 	                IF XGUCODIGO <> "NO" THEN
 	                   XGUFNC2 = " document.all['---" & XGUCODIGO & "'].scrollIntoView(true); " & chr(13)
 	                ELSE
 	                   XGUFNC2 = ""
 	                END IF



 	                    XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
 	                    XGUFBTD1 = XGUFBTD1 & " >"

 	                    XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECIONAR FILTRO.'  "

 	                    IF XGUITEMCD <> "NO" THEN
 	                       XGUFBTD1 = XGUFBTD1 & XGUSTYLS
 	                    ELSE
 	                       XGUFBTD1 = XGUFBTD1 & XGUSTYLN
 	                    END IF

 	                    XGUFBTD1 = XGUFBTD1 & " >"
 	                    XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE=""" & XGUITEMCD& """ CLASS=DT size=40 >&nbsp;&nbsp;"

 	                    XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
	                    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
 	                    XGUFBTD1 = XGUFBTD1 & "</TD>"

 	                    XGUFBTD1 = XGUFBTD1 & "</TR>"


 	                IF XGUITEMCD <> "NO" THEN

 	                   CALL getGruposMail("*" & XGUITEMCD & "*")

 	                   XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
 	                   XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

 	                   xTrColor = "#FFFFFF"

 	                   FOR XGUINX=0 TO UBOUND(XGUTABKEY)


 	                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
 	                   XGUFBTD2 = XGUFBTD2 & " >"

 	                   XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECIONAR GRUPO.' ID='---" & XGUTABKEY(XGUINX) & "' "

 	                   IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
 	                      XGUFBTD2 = XGUFBTD2 & XGUSTYLS
 	                   ELSE
 	                      XGUFBTD2 = XGUFBTD2 & XGUSTYLN
 	                   END IF

	                   XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & replace(XGUITEMCD &"","'","\'")  & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
 	                   XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
 	                   XGUFBTD2 = XGUFBTD2 & "</TR>"


 	                            if xTrColor = "#FFFFFF" then
 	                               xTrColor = "#F9F9F9"
 	                            ELSE
 	                               xTrColor = "#FFFFFF"
 	                            END IF


 	                   NEXT




 	              END IF

 	      CASE ELSE

 	                 XGUFILDSC1 = "&nbsp;"

 	                 XGUFILTRO = "NO"
 	                 XGUITEMCD = "NO"
 	                 XGUCODIGO = "NO"
 	              XGUCODDSC = "NO"

 	  END SELECT

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


  CASE "PORSECPERF"

                  XGUFILDSC = "&nbsp;SECTORES/PERFIL"

                  xTrColor = "#FFFFFF"
                    'RESPONSE.WRITE XGUITEMCD & "22"
                  IF XGUITEMCD <> "NO" THEN
                      XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
                  ELSE
                      XGUFNC1 = ""
                  END IF


                  IF XGUCODIGO <> "NO" THEN
                     XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
                  ELSE
                     XGUFNC2 = ""
                  END IF



                 ' strsql = "Select e.* From Eco e "
                 ' strsql =strsql & "where e.eco in(select distinct l.codsector from objsector l where l.codobjeto=0"& session("NCliente") &")"
                 ' strsql =strsql & " order by e.descri"
                 strsql= " Select A.OBJID,B.DESCRI From objsector A "
				 strsql=strsql & " INNER JOIN ECO B ON(B.ECO= A.CODSECTOR)"
                 strsql=strsql & " WHERE A.CODOBJETO=0" & session("NCliente")
                  set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

                  do until xgurs.eof


                     XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                     XGUFBTD1 = XGUFBTD1 & " >"

                     XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

                     IF CSTR(xgurs(0)) = XGUITEMCD THEN
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                     ELSE
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                     END IF

                     XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                     XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                     XGUFBTD1 = XGUFBTD1 & "</TR>"


                              if xTrColor = "#FFFFFF" then
                                 xTrColor = "#F9F9F9"
                              ELSE
                                 xTrColor = "#FFFFFF"
                              END IF


                  xgurs.movenext
                  loop

                  call ofv.cerrarconsulta(xgurs)

                  session("SGU_UIGuserid")=""
                  session("SGU_UIGdescri")=""
                  IF XGUITEMCD <> "NO" THEN

                    ' CALL getUsrInSECT(XGUITEMCD,"")
                    if XGUITEMCD=0 then XGUITEMCD=-1
                     'sql = "Select Ecoperfil.*,Tipoperfil.descri From Ecoperfil, Tipoperfil where "
                     'sql = sql & "Ecoperfil.codperfil = Tipoperfil.tipoperfil and codsector=" & XGUITEMCD

                      SQL= "Select A.*,B.descri From Ecoperfil A"
					  SQL=SQL & " INNER JOIN Tipoperfil B ON( A.codperfil=B.tipoperfil)"
					  SQL=SQL & " INNER JOIN  objsector C ON (C.OBJID=0"&XGUITEMCD&" AND C.CODSECTOR=A.CODSECTOR)"

                    set xpuestos = ofv.crearconsultaEx(Sql,xgucn,1,volver)
                     delim=""

                     do while not xpuestos.eof

                     session("SGU_UIGuserid")= session("SGU_UIGuserid")   & delim  & xpuestos(0)
                     session("SGU_UIGdescri")= session("SGU_UIGdescri")  & delim  & xpuestos(4)
                     delim =chr(9)
                     xpuestos.movenext
                     loop
                     call  ofv.cerrarconsulta(xpuestos)
                     XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                     XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                     xTrColor = "#FFFFFF"

                     FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                      XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                     XGUFBTD2 = XGUFBTD2 & " >"

                     XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR PUESTO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                     IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                        XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                     ELSE
                        XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                     END IF

                     XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                     XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                     XGUFBTD2 = XGUFBTD2 & "</TR>"


                              if xTrColor = "#FFFFFF" then
                                 xTrColor = "#F9F9F9"
                              ELSE
                                 xTrColor = "#FFFFFF"
                              END IF


                     NEXT




                END IF
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



  CASE "PORPUESTO"

   XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
   XGUFBBOT2  = ""
   XGUFBBOT3  = ""
   XGUFBBOT4  = ""


              XGUFILDSC = "&nbsp;POR&nbsp;PUESTO"

              xTrColor = "#FFFFFF"

              IF XGUCODIGO <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF

              strsql = "SELECT B.OBJID,B.DESCRI + ' - ' + C.DESCRI collate database_default AS T1 FROM LISTADEPUESTOS B  INNER JOIN ECO C ON (C.ECO=B.SECTOR )  ORDER BY T1  "
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR PUESTO.'  ID='***" & xgurs(0) & "' "

                 IF (XGUM3FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF



                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM3FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)


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




       CASE "E"

  XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
  XGUFBBOT2  = ""
  XGUFBBOT3  = ""
  XGUFBBOT4  = ""

              XGUFILDSC = "&nbsp;EMPRESA"

              xTrColor = "#FFFFFF"

              IF XGUCODIGO <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF

              strsql = "Select * From clienteplan  order by cliente"
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR EMPRESA.' ID='***" & xgurs(0) & "' "

                 IF (XGUM1FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF



                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM1FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)


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




       CASE "S"

  XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
  XGUFBBOT2  = ""
  XGUFBBOT3  = ""
  XGUFBBOT4  = ""

              XGUFILDSC = "&nbsp;SECTOR"

              xTrColor = "#FFFFFF"

              IF XGUCODIGO <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF



              strsql    = "Select * From Eco  order by descri"
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR SECTOR.'  ID='***" & xgurs(0) & "' "

                 IF (XGUM2FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF



                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM2FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)


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


       CASE "G"

  XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
  XGUFBBOT2  = ""
  XGUFBBOT3  = ""
  XGUFBBOT4  = ""


              XGUFILDSC = "&nbsp;GRUPO"

              xTrColor = "#FFFFFF"

              IF XGUCODIGO <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF

             ' strsql = "Select * From GRUPOS  order by descriPCION "
             if session("USRPRV")  = "I" or session("USRPRV")  = "E" then
				   strsql = "Select * From GRUPOS  order by descriPCION "
				  else
				   strsql = "Select * From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
             end if
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR GRUPO.'  ID='***" & xgurs(0) & "' "

                 IF (XGUM3FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF



                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM3FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)


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






       CASE "U"




  XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
  XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
  XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
  XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

  XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
  XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
  XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
  XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"



         XGUFILDSC = "&nbsp;POR&nbsp;USUARIO"

  SELECT CASE XGUFILTRO

         CASE "G"

              XGUFILDSC1 = ":&nbsp;GRUPO"

              xTrColor = "#FFFFFF"

              IF XGUITEMCD <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF


              IF XGUCODIGO <> "NO" THEN
                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                 XGUFNC2 = ""
              END IF


              'strsql = "Select * From GRUPOS  order by descriPCION "
              if session("USRPRV")  = "I" or session("USRPRV")  = "E" then
				strsql = "Select * From GRUPOS  order by descriPCION "
			   else
				strsql = "Select * From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
             end if
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.' ID='***" & xgurs(0) & "' "

                 IF CSTR(xgurs(0)) = XGUITEMCD THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF



                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)

              IF XGUITEMCD <> "NO" THEN

                 CALL getUsrIngroup(XGUITEMCD)

                 XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                 XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                 xTrColor = "#FFFFFF"

                 WWWW = INSTR(UCASE(REQUEST.SERVERVARIABLES("SCRIPT_NAME") & ""),"AGREGARDIST.ASP")

                 IF WWWW > 0 THEN

 				                  XGUFBTD2 = XGUFBTD2 & "<TR bgcolor='#6699CC' "
 				                  XGUFBTD2 = XGUFBTD2 & " >"

 				                  XGUFBTD2 = XGUFBTD2 & "<TD CLASS=VTI STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#000000;COLOR:#FFFFFF;' "
 				                  XGUFBTD2 = XGUFBTD2 & " TITLE='SELECCIONAR TODOS LOS USUARIOS.'  ID='---XXXTODOSXXX" & XGUITEMCD & "' "

 				                  IF (XGUM4FL & "TODOS") = XGUCODIGO THEN
 				                     XGUFBTD2 = XGUFBTD2 & XGUSTYLS
 				                  ELSE
 				                     XGUFBTD2 = XGUFBTD2 & XGUSTYLN
 				                  END IF

 				                  XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & "XXXTODOSXXX" & XGUITEMCD & "','" & "TODOS LOS USUARIOS" & "');document.getElementById('XGUFRM').submit();return false;"" >"
 				                  XGUFBTD2 = XGUFBTD2 & "TODOS LOS USUARIOS" & "</TD>"
 				                  XGUFBTD2 = XGUFBTD2 & "</TR>"

               END IF


                 FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                  XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD2 = XGUFBTD2 & " >"

                 XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

                 IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                 ELSE
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                 END IF

                 XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                 XGUFBTD2 = XGUFBTD2 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


                 NEXT




              END IF



         CASE "S"

              XGUFILDSC1 = ":&nbsp;SECTOR"

              xTrColor = "#FFFFFF"

              IF XGUITEMCD <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF


              IF XGUCODIGO <> "NO" THEN
                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                 XGUFNC2 = ""
              END IF



              strsql = "Select * From Eco  order by descri"
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

                 IF CSTR(xgurs(0)) = XGUITEMCD THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF

                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)


              IF XGUITEMCD <> "NO" THEN

                 CALL getUsrInSECT(XGUITEMCD,"")

                 XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                 XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                 xTrColor = "#FFFFFF"

                 FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                  XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD2 = XGUFBTD2 & " >"

                 XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                 IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                 ELSE
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                 END IF

                 XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                 XGUFBTD2 = XGUFBTD2 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


                 NEXT




              END IF




         CASE "L"

              XGUFILDSC1 = ":&nbsp;LETRA"

              xTrColor = "#F9F9F9"


              IF XGUCODIGO <> "NO" THEN
                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                 XGUFNC2 = ""
              END IF


              dim ATABXGULTR(28)

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


              FOR XGUINX=1 TO UBOUND(ATABXGULTR) STEP 4

                  XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                  XGUFBTD1 = XGUFBTD1 & " >"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  IF ATABXGULTR(XGUINX) = XGUITEMCD THEN
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                  ELSE
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                  END IF

                  XGUFBTD1 = XGUFBTD1 & " >"

                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
                  XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                  XGUFBTD1 = XGUFBTD1 & "</TD>"


                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  IF ATABXGULTR(XGUINX+1) = XGUITEMCD THEN
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                  ELSE
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                  END IF

                  XGUFBTD1 = XGUFBTD1 & " >"
                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
                  XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                  XGUFBTD1 = XGUFBTD1 & "</TD>"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  if  ATABXGULTR(XGUINX+2) = "NO" then

                     XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                  else
                     IF ATABXGULTR(XGUINX+2) = XGUITEMCD THEN
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                     ELSE
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                     END IF
                     XGUFBTD1 = XGUFBTD1 & " >"
                     XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
                     XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

                  end if

                  XGUFBTD1 = XGUFBTD1 & "</TD>"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  if  ATABXGULTR(XGUINX+3) = "NO" then

                     XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                  else
                     IF ATABXGULTR(XGUINX+3) = XGUITEMCD THEN
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                     ELSE
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                     END IF
                     XGUFBTD1 = XGUFBTD1 & " >"
                     XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
                     XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

                  end if

                  XGUFBTD1 = XGUFBTD1 & "</TD>"



                  XGUFBTD1 = XGUFBTD1 & "</TR>"


                  if xTrColor = "#FFFFFF" then
                     xTrColor = "#F9F9F9"
                  ELSE
                     xTrColor = "#FFFFFF"
                  END IF

              NEXT




              IF XGUITEMCD <> "NO" THEN

                 CALL getUsrInTEXT2(XGUITEMCD & "*")

                 XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                 XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                 xTrColor = "#FFFFFF"

                 FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                  XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD2 = XGUFBTD2 & " >"

                 XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                 IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                 ELSE
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                 END IF

                 XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                 XGUFBTD2 = XGUFBTD2 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


                 NEXT




              END IF




         CASE "T"

              XGUFILDSC1 = ":&nbsp;TEXTO"



              xTrColor = "#FFFFFF"


              IF XGUCODIGO <> "NO" THEN
                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                 XGUFNC2 = ""
              END IF



                  XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                  XGUFBTD1 = XGUFBTD1 & " >"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

                  IF XGUITEMCD <> "NO" THEN
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                  ELSE
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                  END IF

                  XGUFBTD1 = XGUFBTD1 & " >"
                  XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE=""" & XGUITEMCD & """ CLASS=DT size=40 >&nbsp;&nbsp"

                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
                  XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                  XGUFBTD1 = XGUFBTD1 & "</TD>"

                  XGUFBTD1 = XGUFBTD1 & "</TR>"


              IF XGUITEMCD <> "NO" THEN

                 CALL getUsrInTEXT2("*" & XGUITEMCD & "*")

                 XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                 XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                 xTrColor = "#FFFFFF"

                 FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                  XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD2 = XGUFBTD2 & " >"

                 XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                 IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                 ELSE
                    XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                 END IF

                 XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" &  replace(XGUITEMCD &"","'","\'") & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                 XGUFBTD2 = XGUFBTD2 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


                 NEXT




              END IF






         CASE ELSE

              XGUFILDSC1 = "&nbsp;"

              XGUFILTRO = "NO"
              XGUITEMCD = "NO"
              XGUCODIGO = "NO"
              XGUCODDSC = "NO"


  END SELECT






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


  END SELECT


   call ofv.cerrarconn(xgucn)


  Whtml = ""

  Whtml = Whtml & "<center>"

  if XGUMASTER="DEPENDENCIA" OR XGUMASTER="PORPUESTO" THEN
  XANCHO="95%"
  ELSE
  XANCHO="60%"
  END IF
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

  if XGUCODDSC <> "NO" then
     Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
  else
     Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
  end if

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

  IF XGUMASTERW="PORSECPERF" OR  XGUMASTERW="PORSECTOREX" THEN
  Whtml = Whtml & "<FORM NAME='XGUFRM' id='XGUFRM' TARGET=WK4 ACTION='" & XGUFBURL & "' METHOD='POST' >"
  ELSE
  Whtml = Whtml & "<FORM NAME='XGUFRM'  id='XGUFRM' ACTION='" & XGUFBURL & "' METHOD='POST' >"
  END IF
  Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUMASTER' id='XGUMASTER' >"
  Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUFILTRO' id='XGUFILTRO'>"
  Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUITEMCD' id='XGUITEMCD'>"
  Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODIGO' id='XGUCODIGO'>"
  Whtml = Whtml & "<INPUT TYPE='HIDDEN' VALUE='NO' NAME='XGUCODDSC' id='XGUCODDSC'>"
  Whtml = Whtml & "</FORM>"

  Whtml = Whtml & "<script language=javascript >"
  Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4) { " & chr(13)
  Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & chr(13)
  Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU1; " & chr(13)
  Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU2;" & chr(13)
'Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=document.getElementById('XGUITEMCD').value.replace('\'','fxx');" & chr(13)
  Whtml = Whtml & "    document.getElementById('XGUCODIGO').value=XGU3; " & chr(13)
  Whtml = Whtml & "    document.getElementById('XGUCODDSC').value=XGU4; " & chr(13)
  Whtml = Whtml & " } " & chr(13)
  Whtml = Whtml & XGUFNC1
  Whtml = Whtml & XGUFNC2
  Whtml = Whtml & "</script>"



  SESSION("FBUSR") = XGUCODIGO
  SESSION("FBSLC") = XGUMASTER
  session("XGUCODDSC")=XGUCODDSC
SESSION("XGUITEMCD")=XGUITEMCD

      GenBuscaUsrEx = Whtml




 end function


 function GenBuscaUsr(XGUM1,XGUM2,XGUM3,XGUM4,XGUM1F,XGUM2F,XGUM3F,XGUM4F)


 XGUM1FL = REPLACE(XGUM1F,CHR(9),"")
 XGUM2FL = REPLACE(XGUM2F,CHR(9),"")
 XGUM3FL = REPLACE(XGUM3F,CHR(9),"")
 XGUM4FL = REPLACE(XGUM4F,CHR(9),"")


 XGUFBURL = request.servervariables("SCRIPT_NAME")

 XGUFBURL = RIGHT(XGUFBURL,LEN(XGUFBURL) - (INSTRREV(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

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

 IF SESSION("EMPMUN") = "S" then

 IF XGUM1 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "E"
    END IF

    XGUMSBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='EMPRESAS'  CLASS=BT ONCLICK=""clactfrm('E','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
    XGUMSBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI1 & "</FORM>"

 END IF

 END IF

 IF XGUM2 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "S"
    END IF

    XGUMSBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTORES'  CLASS=BT ONCLICK=""clactfrm('S','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
    XGUMSBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI2 & "</FORM>"


 END IF

 IF XGUM3 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "G"
    END IF

    XGUMSBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPOS'  CLASS=BT ONCLICK=""clactfrm('G','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
    XGUMSBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTI3 & "</FORM>"


 END IF

 IF XGUM4 = "S" THEN

    IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTERW = "U"
    END IF

    XGUMSBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='USUARIOS'  CLASS=BT ONCLICK=""clactfrm('U','NO','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
    XGUMSBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUMSBOTi4 & "</FORM>"


 END IF


 IF  XGUMASTER = "" OR XGUMASTER = "NO" THEN
        XGUMASTER = XGUMASTERW
 END IF


 XGUCONTEMS = ""

 XGUCONTEMS = XGUCONTEMS & "<table id=""MTBSUSR0B""  class='aaa'  cellspacing=1 cellPADDING=0 > "
 XGUCONTEMS = XGUCONTEMS & "<tr>"
 XGUCONTEMS = XGUCONTEMS & "<td class=ut WIDTH=60%  >&nbsp;</td>"

 IF XGUMSBOT1 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT1
 END IF

 IF XGUMSBOT2 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT2
 END IF

 IF XGUMSBOT3 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT3
 END IF

 IF XGUMSBOT4 <> "&nbsp;" THEN
    XGUCONTEMS = XGUCONTEMS & XGUMSBOT4
 END IF

 XGUCONTEMS = XGUCONTEMS & "</tr>"
 XGUCONTEMS = XGUCONTEMS & "</table>"


 XGUFNC1 = ""
 XGUFNC2 = ""


 set xgucn = ofv.conectar(ofv.strconn2)

 XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#F7AA22;cursor:pointer;' "
 XGUSTYLN = " STYLE='cursor:pointer;' "

 XGUFBUSR = XGUCODIGO

 XGUFBTAB1  = ""
 XGUFBTAB2  = ""
 XGUFBTD1  = ""
 XGUFBTD2  = ""

 IF XGUMASTER = "" OR XGUMASTER = "NO" THEN XGUMASTER = "U"

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '' SELECCION DE BUSQUEDA MASTER - E:EMPRESAS / S:SECTORES / G:GRUPOS / U:USUARIOS              ''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SESSION ("XGUMASTER" )=XGUMASTER
 SELECT CASE XGUMASTER

      CASE "E"

 XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
 XGUFBBOT2  = ""
 XGUFBBOT3  = ""
 XGUFBBOT4  = ""

             XGUFILDSC = "&nbsp;EMPRESA"

             xTrColor = "#FFFFFF"

             IF XGUCODIGO <> "NO" THEN
                 XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF

             strsql = "Select * From clienteplan  order by cliente"
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR EMPRESA.' ID='***" & xgurs(0) & "' "

                IF (XGUM1FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM1FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


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




      CASE "S"

 XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
 XGUFBBOT2  = ""
 XGUFBBOT3  = ""
 XGUFBBOT4  = ""

             XGUFILDSC = "&nbsp;SECTOR"

             xTrColor = "#FFFFFF"

             IF XGUCODIGO <> "NO" THEN
                 XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF



             strsql    = "Select * From Eco  order by descri"
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR SECTOR.'  ID='***" & xgurs(0) & "' "

                IF (XGUM2FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM2FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


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


      CASE "G"

 XGUFBBOT1  = "<td class=ut  >&nbsp;</td>"
 XGUFBBOT2  = ""
 XGUFBBOT3  = ""
 XGUFBBOT4  = ""


             XGUFILDSC = "&nbsp;GRUPO"

             xTrColor = "#FFFFFF"

             IF XGUCODIGO <> "NO" THEN
                 XGUFNC1 = " document.getElementById('***" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF

            ' strsql = "Select * From GRUPOS  order by descriPCION "
            if session("USRPRV")  = "I" or session("USRPRV")  = "E" then
			  strsql = "Select * From GRUPOS  order by descriPCION "
			 else
			  strsql = "Select * From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
             end if

             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECIONAR GRUPO.'  ID='***" & xgurs(0) & "' "

                IF (XGUM3FL & CSTR(xgurs(0))) = XGUCODIGO THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','NO','"  & XGUM3FL & xgurs(0) & "','" & xgurs(1) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


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






      CASE "U"



 XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='GRUPO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','G','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
 XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='SECTOR' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','S','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
 XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','L','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
 XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','T','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"



 XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
 XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
 XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
 XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"



        XGUFILDSC = "&nbsp;USUARIO"

 SELECT CASE XGUFILTRO

        CASE "G"

             XGUFILDSC1 = ":&nbsp;GRUPO"

             xTrColor = "#FFFFFF"

             IF XGUITEMCD <> "NO" THEN
                 XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF

             if session("USRPRV")  = "I" or session("USRPRV")  = "E" then
              strsql = "Select * From GRUPOS  order by descriPCION "
             else
              strsql = "Select * From GRUPOS where rtrim(ltrim(mapgrp + '')) <> ''  order by descriPCION "
             end if

             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.' ID='***" & xgurs(0) & "' "

                IF CSTR(xgurs(0)) = XGUITEMCD THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF



                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)

             IF XGUITEMCD <> "NO" THEN

                CALL getUsrIngroup(XGUITEMCD)

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                xTrColor = "#FFFFFF"

                WWWW = INSTR(UCASE(REQUEST.SERVERVARIABLES("SCRIPT_NAME") & ""),"AGREGARDIST.ASP")

                IF WWWW > 0 THEN

				                  XGUFBTD2 = XGUFBTD2 & "<TR bgcolor='#6699CC' "
				                  XGUFBTD2 = XGUFBTD2 & " >"

				                  XGUFBTD2 = XGUFBTD2 & "<TD CLASS=VTI STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#000000;COLOR:#FFFFFF;' "
				                  XGUFBTD2 = XGUFBTD2 & " TITLE='SELECCIONAR TODOS LOS USUARIOS.'  ID='---XXXTODOSXXX" & XGUITEMCD & "' "

				                  IF (XGUM4FL & "TODOS") = XGUCODIGO THEN
				                     XGUFBTD2 = XGUFBTD2 & XGUSTYLS
				                  ELSE
				                     XGUFBTD2 = XGUFBTD2 & XGUSTYLN
				                  END IF

				                  XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & "XXXTODOSXXX" & XGUITEMCD & "','" & "TODOS LOS USUARIOS" & "');document.getElementById('XGUFRM').submit();return false;"" >"
				                  XGUFBTD2 = XGUFBTD2 & "TODOS LOS USUARIOS" & "</TD>"
				                  XGUFBTD2 = XGUFBTD2 & "</TR>"

              END IF


                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT




             END IF



        CASE "S"

             XGUFILDSC1 = ":&nbsp;SECTOR"

             xTrColor = "#FFFFFF"

             IF XGUITEMCD <> "NO" THEN
                 XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
             ELSE
                 XGUFNC1 = ""
             END IF


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF



             strsql = "Select * From Eco  order by descri"
             set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

             do until xgurs.eof


                XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD1 = XGUFBTD1 & " >"

                XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='SELECCIONAR FILTRO.'  ID='***" & xgurs(0) & "' "

                IF CSTR(xgurs(0)) = XGUITEMCD THEN
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                ELSE
                   XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                END IF

                XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                XGUFBTD1 = XGUFBTD1 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


             xgurs.movenext
             loop

             call ofv.cerrarconsulta(xgurs)


             IF XGUITEMCD <> "NO" THEN

                CALL getUsrInSECT(XGUITEMCD,"")

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                xTrColor = "#FFFFFF"

                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT




             END IF




        CASE "L"

             XGUFILDSC1 = ":&nbsp;LETRA"

             xTrColor = "#F9F9F9"


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF


             dim ATABXGULTR(28)

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


             FOR XGUINX=1 TO UBOUND(ATABXGULTR) STEP 4

                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                 IF ATABXGULTR(XGUINX) = XGUITEMCD THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF

                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & "</TD>"


                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                 IF ATABXGULTR(XGUINX+1) = XGUITEMCD THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF

                 XGUFBTD1 = XGUFBTD1 & " >"
                 XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & "</TD>"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                 if  ATABXGULTR(XGUINX+2) = "NO" then

                    XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                 else
                    IF ATABXGULTR(XGUINX+2) = XGUITEMCD THEN
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                    ELSE
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                    END IF
                    XGUFBTD1 = XGUFBTD1 & " >"
                    XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
                    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

                 end if

                 XGUFBTD1 = XGUFBTD1 & "</TD>"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                 if  ATABXGULTR(XGUINX+3) = "NO" then

                    XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                 else
                    IF ATABXGULTR(XGUINX+3) = XGUITEMCD THEN
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                    ELSE
                       XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                    END IF
                    XGUFBTD1 = XGUFBTD1 & " >"
                    XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
                    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"

                 end if

                 XGUFBTD1 = XGUFBTD1 & "</TD>"



                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                 if xTrColor = "#FFFFFF" then
                    xTrColor = "#F9F9F9"
                 ELSE
                    xTrColor = "#FFFFFF"
                 END IF

             NEXT




             IF XGUITEMCD <> "NO" THEN

                CALL getUsrInTEXT(XGUITEMCD & "*")

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                xTrColor = "#FFFFFF"

                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT




             END IF




        CASE "T"

             XGUFILDSC1 = ":&nbsp;TEXTO"



             xTrColor = "#FFFFFF"


             IF XGUCODIGO <> "NO" THEN
                XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
             ELSE
                XGUFNC2 = ""
             END IF



                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

                 IF XGUITEMCD <> "NO" THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF

                 XGUFBTD1 = XGUFBTD1 & " >"
                 XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE='" & XGUITEMCD & "' CLASS=DT size=40 >&nbsp;&nbsp;"

                 XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
                 XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1 & "</TD>"

                 XGUFBTD1 = XGUFBTD1 & "</TR>"


             IF XGUITEMCD <> "NO" THEN

                CALL getUsrInTEXT("*" & XGUITEMCD & "*")

                XGUTABKEY = SPLIT(session("SGU_UIGuserid"),CHR(9))
                XGUTABDAT = SPLIT(session("SGU_UIGdescri"),CHR(9))

                xTrColor = "#FFFFFF"

                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                 XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                XGUFBTD2 = XGUFBTD2 & " >"

                XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR USUARIO.' ID='---" & XGUTABKEY(XGUINX) & "' "

                IF (XGUM4FL & XGUTABKEY(XGUINX)) = XGUCODIGO THEN
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                ELSE
                   XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                END IF

                XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & XGUITEMCD & "','" & XGUM4FL & XGUTABKEY(XGUINX) & "','" & XGUTABDAT(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT




             END IF


        CASE ELSE

             XGUFILDSC1 = "&nbsp;"

             XGUFILTRO = "NO"
             XGUITEMCD = "NO"
             XGUCODIGO = "NO"
             XGUCODDSC = "NO"


 END SELECT






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


 END SELECT


  call ofv.cerrarconn(xgucn)


 Whtml = ""

 Whtml = Whtml & "<center>"
 Whtml = Whtml & "<DIV id='BSUSR' class='DINDM'  STYLE='display:block;z-index:99;margin-top:1;margin-left:1;width:650;"
 Whtml = Whtml & "'>"
 Whtml = Whtml & "<CENTER>"

 Whtml = Whtml & "<table id=""MTBSUSR1""  class='TWIND' width=99% cellspacing=0 cellpadding=2 >"
 Whtml = Whtml & "<tr HEIGHT=25 >"
 Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;&nbsp;BUSCAR" & XGUFILDSC & "&nbsp;</td>"
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

 if XGUCODDSC <> "NO" then
    Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
 else
    Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
 end if

 Whtml = Whtml & "</tr>"
 Whtml = Whtml & "</table>"

 Whtml = Whtml & "<table id=""MTBSUSR3""  class='aaa' width=99% cellspacing=1 cellpadding=1 "
 Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 Whtml = Whtml & " >"

 Whtml = Whtml & XGUCONTENT

 Whtml = Whtml & "</table>"
 Whtml = Whtml & "</CENTER>"
 Whtml = Whtml & "</DIV>"
 Whtml = Whtml & "</center>"




 Whtml = Whtml & "<script language=javascript >"
 Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4) { " & chr(13)
 ' Whtml = Whtml & " alert(document.getElementById('XGUFRM'));  "
 Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU1; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU2; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUCODIGO').value=XGU3; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUCODDSC').value=XGU4; " & chr(13)
 Whtml = Whtml & " } " & chr(13)
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



     GenBuscaUsr = Whtml


 end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' FUNCIONES DE BUSQUEDa de Documentos ( documento-solapa-1er nivel)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function GenBuscaObj(XGUCOBJI,XGUTOBJI,XGUN1I,XGUN2I,XGUN3I,XGUN4I,XGUN5I)


 XGUFBURL = request.servervariables("SCRIPT_NAME")

 XGUFBURL = RIGHT(XGUFBURL,LEN(XGUFBURL) - (INSTRREV(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

 XGUMASTER = REQUEST.FORM("XGUMASTER")

  XGUITEMCD = REQUEST.FORM("XGUITEMCD")
  XGUFILTRO = REQUEST.FORM("XGUFILTRO")

'RESPONSE.WRITE XGUCOBJI & "  " &XGUTOBJI& "  " &XGUN1I& "  " &XGUN2I& "  " &XGUN3I& "  " &XGUN4I& "  " &XGUN5I
xcodf=request.querystring("codobjeto")&""



IF  XGUMASTER = "" THEN
     if XGUCOBJI & "" = xcodf then
     XGUFILTRO="11"
     XGUITEMCD = xcodf
     else

	 XGUFILTRO = "TX"
	 XGUITEMCD = "NO"
	 end if



     XGUMASTER = "0"

     XGUFCOBJ = "0"
     XGUFTOBJ = "0"
     XGUFNIV1 = "0"
     XGUFNIV2 = "0"
     XGUFNIV3 = "0"
     XGUFNIV4 = "0"
     XGUFNIV5 = "0"

     if   XGUCOBJI <> "0" THEN
          XGUFCOBJ = XGUCOBJI
          XGUFTOBJ = XGUTOBJI
     END IF

     if   XGUN1I <> "0" THEN
          XGUFNIV1 = XGUN1I
          XGUMASTER = "1"
     END IF

     if   XGUN2I <> "0" THEN
          XGUFNIV2 = XGUN2I
          XGUMASTER = "2"
     END IF

     if   XGUN3I <> "0" THEN
          XGUFNIV3 = XGUN3I
          XGUMASTER = "3"
     END IF

     if   XGUN4I <> "0" THEN
          XGUFNIV4 = XGUN4I
          XGUMASTER = "4"
     END IF

     if   XGUN5I <> "0" THEN
          XGUFNIV5 = XGUN5I
          XGUMASTER = "5"
     END IF

ELSE

     XGUFCOBJ = REQUEST.FORM("XGUFCOBJ")
     XGUFTOBJ = REQUEST.FORM("XGUFTOBJ")
     XGUFNIV1 = REQUEST.FORM("XGUFNIV1")
     XGUFNIV2 = REQUEST.FORM("XGUFNIV2")
     XGUFNIV3 = REQUEST.FORM("XGUFNIV3")
     XGUFNIV4 = REQUEST.FORM("XGUFNIV4")
     XGUFNIV5 = REQUEST.FORM("XGUFNIV5")


 END IF



 XGUCONTEMS = ""



 XGUFNC1 = ""
 XGUFNC2 = ""


 set xgucn = ofv.conectar(ofv.strconn2)

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
  if XGUMASTER <> "11" AND XGUFILTRO <> "11"   then
   XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TIPO' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TI','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
   XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TL','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
   XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TX','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

   XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
   XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
   XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
   XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"

 end if

          XGUFILDSC = "&nbsp;componentes"
'RESPONSE.WRITE "<BR>eeeeeeeeeeeeeeeeeeeeeeee" & XGUFILTRO  & "<br>eeeeeeeeeeeeee "  & XGUITEMCD
   SELECT CASE XGUFILTRO

          CASE "NO"

          CASE "TI"

               XGUFILDSC1 = ":&nbsp;TIPO COMPONENTE"

               xTrColor = "#FFFFFF"

               IF XGUITEMCD <> "NO" THEN
                   XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
               ELSE
                   XGUFNC1 = ""
               END IF


               IF XGUCODIGO <> "NO" THEN
                  XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
               ELSE
                  XGUFNC2 = ""
               END IF


               strsql = "Select * From TIPOBJETO  where aplic IN ('EM', 'GLB') order by descri "
             ''response.write strsql & " strsql " & "<br>"
               set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

               do until xgurs.eof


                  XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                  XGUFBTD1 = XGUFBTD1 & " >"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='Seleccionar item " & xgurs(1) & ".' ID='***" & xgurs(0) & "' "

                  IF CSTR(xgurs(0)) = XGUITEMCD THEN
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                  ELSE
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                  END IF



           '''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                  XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TI','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TI','" & xgurs(0) & "');document.getElementById('XGUFRM').submit();return false;"" >"


                  XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                  XGUFBTD1 = XGUFBTD1 & "</TR>"


                           if xTrColor = "#FFFFFF" then
                              xTrColor = "#F9F9F9"
                           ELSE
                              xTrColor = "#FFFFFF"
                           END IF


               xgurs.movenext
               loop

               call ofv.cerrarconsulta(xgurs)


              XGUFBTD1xxx = XGUFBTD1






          CASE "TL"

               XGUFILDSC1 = ":&nbsp;LETRA"

               xTrColor = "#F9F9F9"


               IF XGUCODIGO <> "NO" THEN
                  XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
               ELSE
                  XGUFNC2 = ""
               END IF


               dim ATABXGULTR(28)

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


               FOR XGUINX=1 TO UBOUND(ATABXGULTR) STEP 4

                   XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                   XGUFBTD1 = XGUFBTD1 & " >"

                   XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                   IF ATABXGULTR(XGUINX) = XGUITEMCD THEN
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                   ELSE
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                   END IF

                   XGUFBTD1 = XGUFBTD1 & " >"

                   XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
             ''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


                   XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"





                   XGUFBTD1 = XGUFBTD1 & "</TD>"


                   XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                   IF ATABXGULTR(XGUINX+1) = XGUITEMCD THEN
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                   ELSE
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                   END IF

                   XGUFBTD1 = XGUFBTD1 & " >"
                   XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
             ''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



                   XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX+1) & "');document.getElementById('XGUFRM').submit();return false;"" >"





                   XGUFBTD1 = XGUFBTD1 & "</TD>"

                   XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                   if  ATABXGULTR(XGUINX+2) = "NO" then

                      XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                   else
                      IF ATABXGULTR(XGUINX+2) = XGUITEMCD THEN
                         XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                      ELSE
                         XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                      END IF
                      XGUFBTD1 = XGUFBTD1 & " >"
                      XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
                ''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



                   XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX+2) & "');document.getElementById('XGUFRM').submit();return false;"" >"






                   end if

                   XGUFBTD1 = XGUFBTD1 & "</TD>"

                   XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                   if  ATABXGULTR(XGUINX+3) = "NO" then

                      XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                   else
                      IF ATABXGULTR(XGUINX+3) = XGUITEMCD THEN
                         XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                      ELSE
                         XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                      END IF
                      XGUFBTD1 = XGUFBTD1 & " >"
                      XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
               '''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


                   XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TL','" & ATABXGULTR(XGUINX+3) & "');document.getElementById('XGUFRM').submit();return false;"" >"




                  end if

                   XGUFBTD1 = XGUFBTD1 & "</TD>"



                   XGUFBTD1 = XGUFBTD1 & "</TR>"


                   if xTrColor = "#FFFFFF" then
                      xTrColor = "#F9F9F9"
                   ELSE
                      xTrColor = "#FFFFFF"
                   END IF

               NEXT









          CASE "TX"

               XGUFILDSC1 = ":&nbsp;TEXTO"



               xTrColor = "#FFFFFF"


               IF XGUCODIGO <> "NO" THEN
                  XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
               ELSE
                  XGUFNC2 = ""
               END IF



                   XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                   XGUFBTD1 = XGUFBTD1 & " >"

                   XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

                   IF XGUITEMCD <> "NO" THEN
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                   ELSE
                      XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                   END IF

                   XGUFBTD1 = XGUFBTD1 & " >"
                   XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE='" & XGUITEMCD & "' CLASS=DT size=40 >&nbsp;&nbsp;"

                   XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
               ''    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


                  XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TX','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','TX',document.getElementById('XGUBTEXTO').value);document.getElementById('XGUFRM').submit();return false;"" >"




                   XGUFBTD1 = XGUFBTD1 & "</TD>"

                   XGUFBTD1 = XGUFBTD1 & "</TR>"



         CASE "11"


          CASE ELSE

               XGUFILDSC1 = "&nbsp;"

               XGUFILTRO = "11"
               XGUITEMCD = "11"
               XGUCODIGO = "NO"
               XGUCODDSC = "NO"


   END SELECT






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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ARMADO DE DIV SEGUN NIVEL DE PROFUNDIDAD ...
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 SELECT CASE XGUMASTER



           '  IF XGUCODIGO <> "NO" THEN
           '     XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
           '  ELSE
           '     XGUFNC2 = ""
           '  END IF

 CASE "TI","TL","TX"




	                   CALL GETLISTOBJ(XGUMASTER,"0",XGUITEMCD)


	                   XGUTABKEY  = SPLIT(XGUTABLISKEY,CHR(9))
	                   XGUTABKEY2 = SPLIT(XGUTABLISKEY2,CHR(9))
	                   XGUTABDAT  = SPLIT(XGUTABLISDAT,CHR(9))

	                   xTrColor = "#FFFFFF"



	                   IF XGUFCOBJ <> "0" THEN
	                      XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & chr(13)
	                   ELSE
	                      XGUFNC2 = ""
	                   END IF



	                   FOR XGUINX=0 TO UBOUND(XGUTABKEY)


	                       XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
	                       XGUFBTD2 = XGUFBTD2 & " >"

	                       XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR COMPONENTE.'  ID='---" & XGUTABKEY(XGUINX) & "' "

	                       IF CSTR(XGUTABKEY(XGUINX)) = cstr(XGUFCOBJ) THEN
	                          XGUFBTD2 = XGUFBTD2 & XGUSTYLS
	                       ELSE
	                          XGUFBTD2 = XGUFBTD2 & XGUSTYLN
	                       END IF



	                       XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
	                       XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
	                       XGUFBTD2 = XGUFBTD2 & "</TR>"


	                            if xTrColor = "#FFFFFF" then
	                               xTrColor = "#F9F9F9"
	                            ELSE
	                               xTrColor = "#FFFFFF"
	                            END IF


	                   NEXT
        CASE "0"

               ' CALL GETLISTOBJ("L","0","0")
               CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUTABKEY  = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABKEY2 = SPLIT(XGUTABLISKEY2,CHR(9))
                XGUTABDAT  = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"



                IF XGUFCOBJ <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR DOCUMENTO.'  ID='---" & XGUTABKEY(XGUINX) & "' "

                    IF CSTR(XGUTABKEY(XGUINX)) = cstr(XGUFCOBJ) THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF



                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT




        CASE "1","11"


                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
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


                CALL GETLISTSLP("L",XGUFCOBJ,XGUFTOBJ,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"


                IF XGUFNIV1 <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFNIV1 & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR SOLAPA.'  ID='---" & XGUTABKEY(XGUINX) & "' "


                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV1 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF

                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUTABKEY(XGUINX) & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT



        CASE "2"

                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"





                CALL GETLISTNV2("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"


                IF XGUFNIV2 <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFNIV2 & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF


                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 1ER. NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV2 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUTABKEY(XGUINX) & "','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT





        CASE "3"

                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICKX=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"





                CALL GETLISTNV3("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"

                IF XGUFNIV3 <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFNIV3 & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 2DO.NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV3 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('4','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUTABKEY(XGUINX) & "','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT


        CASE "4"

                CALL GETLISTOBJ("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICKX=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','"&   XGUfiltro &"','"&   XGUITEMCD&"');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"



                CALL GETLISTNV3("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,XGUFNIV3)
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






 END SELECT



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'XGUMASTER

 IF XGUMASTER > "0" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


 END IF

 IF XGUMASTER > "1" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


 END IF

 IF XGUMASTER > "2" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


 END IF

 IF XGUMASTER > "3" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


 END IF

 IF XGUMASTER > "0" THEN

    XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;width:99%;background-color:#ffffff;"
    XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;'>"
    XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

 END IF




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


  call ofv.cerrarconn(xgucn)


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
  Whtml = Whtml & XGUFBBOT1
  Whtml = Whtml & XGUFBBOT2
  Whtml = Whtml & XGUFBBOT3
  Whtml = Whtml & XGUFBBOT4
  Whtml = Whtml & XGUFBBOT5

  if XGUCODDSC <> "NO" then
     Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
  else
     Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
  end if

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
 Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4,XGU5,XGU6,XGU7,XGU8,XGU9) { " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFCOBJ').value=XGU1;  " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFTOBJ').value=XGU2;  " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV1').value=XGU3;  " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV2').value=XGU4;  " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV3').value=XGU5;  " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV4').value=XGU6;  " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV5').value=XGU7;  " & chr(13)

  Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU8; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU9; " & chr(13)


 Whtml = Whtml & " } " & chr(13)
 Whtml = Whtml & XGUFNC1
 Whtml = Whtml & XGUFNC2
 Whtml = Whtml & "</script>"




     GenBuscaObj = Whtml


 end function






FUNCTION GETLISTNV4(XGUFOPC, XGUFCOBJF, XGUFTOBJF, XGUFNIV1F, XGUFNIV2F, XGUFNIV3F, XGUFNIV4F)

   XGUTABLISKEY = "0"
   XGUTABLISDAT = ""





      IF XGUFNIV1F = "9" THEN

         Strsql = "SELECT * FROM ObjTabcon WHERE OBJID = 0" & XGUFNIV4F

         SET Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
         IF NOT Xrs.EOF THEN
            XGUTABLISKEY = Xrs(0)

            xxxdescritmp = Xrs(6) & " " & Xrs(7) & ""
            if len(xxxdescritmp) > 20 then xxxdescritmp = left(xxxdescritmp & "",20) & "..."
            XGUTABLISDAT = Xrs(4) & "." & xxxdescritmp

         END IF
         CALL ofv.cerrarconsulta(Xrs)

      END IF

      IF XGUFNIV1F = "11" THEN

         Strsql = "SELECT * FROM OBJESTDET WHERE OBJID = 0" & XGUFNIV4F
         SET Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
         IF NOT Xrs.EOF THEN
            XGUTABLISKEY = Xrs(0)

            xxxdescritmp = Xrs(8) & ""
            xxxdescrityp = ucase(Xrs(5) & "")

            if xxxdescrityp = "H" then
               xxxdescritmp = "Imagen asociada."
            elseif xxxdescrityp = "I" then
               xxxdescritmp = "Tabla asociada."
            elseif xxxdescrityp = "J" then
               xxxdescritmp = "Componente asociado."
            elseif xxxdescrityp = "G" then
              if len(xxxdescritmp) > 25 then xxxdescritmp = left(xxxdescritmp & "",25) & "..."
            else
              xxxdescritmp = Xrs(7) & ""
            end if

            XGUTABLISDAT = Xrs(4) & "." & xxxdescritmp
         END IF
         CALL ofv.cerrarconsulta(Xrs)

      END IF



   GETLISTNV4 = "OK"

END FUNCTION

 function GETLISTOBJ(XGUFOPC,XGUFCOBJF,XGUFTOBJF)

     XGUTABLISKEY  = "0"
     XGUTABLISKEY2 = "NO"
     XGUTABLISDAT  = "SIN DEFINIR"

     XGUTABSEP    = CHR(9)

     IF XGUFOPC = "E" THEN

     '  Strsql = "Select * From csObjetosRDO where CODOBJETO = 0" & XGUFCOBJF
       IF Session("CorApli") = "COREM20" THEN
	   	  	Strsql = "SELECT * FROM csobjetosRDOAPL WHERE aplic IN ('EM', 'GLB') and  CODOBJETO = 0" & XGUFCOBJF
	   	  else
	         Strsql = "Select * From csObjetosRDO where CODOBJETO = 0" & XGUFCOBJF
	  end if

        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY  = Xrs2(0)
           XGUTABLISKEY2 = Xrs2(8)
           XGUTABLISDAT  = Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"


        END IF
        call ofv.cerrarconsulta(Xrs2)


     ELSEIF XGUFOPC = "TI" THEN

        Strsql = "Select * From csObjetosRDO where estado = 3  OR ESTADO = 1 order by descri, OBJDESCRI"
       ' response.write Strsql
         IF Session("CorApli") = "COREM20" THEN
			  Strsql = "SELECT * FROM csobjetosRDOAPL WHERE aplic IN ('EM', 'GLB') and (estado = 3  OR ESTADO = 1) and tipobjeto='" & XGUFTOBJF & "' order by descri, OBJDESCRI "
		 else
		      Strsql = "Select * From csObjetosRDO where (estado = 3  OR ESTADO = 1) and tipobjeto='" & XGUFTOBJF & "' order by descri, OBJDESCRI"
	     end if

        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        do while not Xrs2.eof

           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(8)
           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"



        Xrs2.movenext
        loop


        ELSEIF XGUFOPC = "TX" THEN

		        Strsql = "Select * From csObjetosRDO where estado = 3  OR ESTADO = 1 order by descri, OBJDESCRI"
		         IF Session("CorApli") = "COREM20" THEN
					  Strsql = "SELECT * FROM csobjetosRDOAPL WHERE aplic IN ('EM', 'GLB') and (estado = 3  OR ESTADO = 1) and OBJDescri like('%" & XGUFTOBJF & "%')  order by descri, OBJDESCRI "
				 else
				      Strsql = "Select * From csObjetosRDO where (estado = 3  OR ESTADO = 1) and OBJDescri like('%" & XGUFTOBJF & "%')   order by descri, OBJDESCRI"
			     end if

		        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		        do while not Xrs2.eof

		           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
		           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(8)
		           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"



		        Xrs2.movenext
        loop

        ELSEIF XGUFOPC = "TL" THEN

				        //Strsql = "Select * From csObjetosRDO where estado = 3  OR ESTADO = 1 order by descri, OBJDESCRI"
				         IF Session("CorApli") = "COREM20" THEN
							  Strsql = "SELECT * FROM csobjetosRDOAPL WHERE aplic IN ('EM', 'GLB') and (estado = 3  OR ESTADO = 1) and OBJDescri like('" & XGUFTOBJF & "%')  order by descri, OBJDESCRI "
						 else
						      Strsql = "Select * From csObjetosRDO where (estado = 3  OR ESTADO = 1) and OBJDescri like('" & XGUFTOBJF & "%')  order by descri, OBJDESCRI"
					     end if

				        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

				        do while not Xrs2.eof

				           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
				           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(8)
				           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(1) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(3) & "&nbsp;(" & Xrs2(5) & "." & Xrs2(6) & ")"



				        Xrs2.movenext
        loop


        call ofv.cerrarconsulta(Xrs2)

     END IF

     GETLISTOBJ = "OK"


 end function



 function GETLISTSLP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL DOCUMENTO"

     XGUTABSEP    = CHR(9)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' BUSCA LAS SOLAPAS DEL DOCUMENTO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  dim stabla(20)
  dim solord(20)
  dim Nombres(20)
  dim Solapas(20)

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


    for xx = 1 to 20
        solord(xx) = 17
        Nombres(xx) = "N"
    next


  Strsql = "Select * From Objperfildoc Where codobjeto =" & XGUFCOBJF & " and tipobjeto ='" & XGUFTOBJF & "' "
 'response.write Strsql & " strsql 1 " & "<br>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Xrs.eof then
     perfilsp = "N"
  end if

  if perfilsp = "N" then
     call ofv.cerrarconsulta(Xrs)

        sPerfil = "OBJMAIN"

     Strsql = "Select PERFIL From TIPOBJETO Where tipoBJETO='" & XGUFTOBJF & "'"
    'response.write Strsql & " strsql 2 " & "<br>"
     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not Xrs.eof then
        sPerfil = XRS(0)
     end if

     Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
    'response.write Strsql & " strsql 3 " & "<br>"
     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if

 ''response.write strsql & "<br>"


    for xx = 2 to 16
        stabla (xx - 1) = Xrs(xx + 30)
    next


    yy = 0
    for xx = 1 to 15

        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           solord (stabla (xx)) = xx
        end if
    next
    tyy = yy



  for z=0 to 20
      Solapas(z) = "N"
  next


  X=0
  for z = 2 to 16

    if Xrs(z) Then

    '''if (z <> 2 and z <> 4 and z <> 5 )  then
       if (z <> 2 and z <> 5 )  then

          x = x + 1
          Solapas(z-1)  = XRS(z + 15)
         'response.write Solapas(z-1) & " " & z-1 & " Solapas (z) " & "<br>"
          if  z = 6 then

             Solapas(16)  = "CIRCUITO COMPLETO"

             for xiord=0 to 19

                 if solord(xiord) = 5 then

                    xmovey = 16

                    for xiord2=xiord + 1 to 19
                        xmoven           = solord(xiord2)
                        solord(xiord2)   = xmovey
                        xmovey           = xmoven
                    next



                    exit for
                 end if
             next


          end if

      end if
    end if

  next


  call ofv.cerrarconsulta(Xrs)

  xfinal = ubound(Solapas)






     IF XGUFOPC = "E" THEN

        Strsql = "Select * From csGenObjetos where CODOBJETO = 0" & XGUFCOBJF
       'response.write Strsql & " strsql " & "<br>"
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY = Xrs2(0)
           XGUTABLISDAT = Xrs2(1)


        END IF
        call ofv.cerrarconsulta(Xrs2)

        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then
            '   response.write solord(x)
               if clng(XGUFNIV1F) = clng(solord(x)) then

                 XGUTABLISKEY = solord(x)
                 XGUTABLISDAT = solapas(solord(x))

                 EXIT FOR

               END IF
             end if
        next



     ELSE


        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then

                 XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & solord(x)
                 XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & solapas(solord(x))

            end if
        next




     END IF

     GETLISTSLP = "OK"

 end function



 function GETLISTNV2(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO "

     XGUTABSEP    = CHR(9)

     XGUTABLAF    = "NO"


     SELECT CASE XGUFNIV1F

            CASE "2"   'CONSIDERACIONES

                 XGUTABLAF    = "OBJNWEDT"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM OBJNWEDT WHERE OBJID = 0" & XGUFNIV2F
                    'response.write Strsql & " strsql 4818 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM OBJNWEDT WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by consecuen "
                    'response.write Strsql & " strsql 4833 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF

'' NUEVO POR VERSION 4.6 - 25-02-2015

            CASE "3"   'SECTORES

                       'response.write " entro a Sectores - linea 4851 " & "<br>"

	                XGUTABLAF    = "OBJSECTOR"

	                IF XGUFOPC = "E" THEN

	                  Strsql = "SELECT OBJSECTOR.*, ECO.descri FROM OBJSECTOR, ECO "
	                  Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and eco.estado=3 and codobjeto='" & XGUFNIV2F & "' "
	                  Strsql = Strsql & " And (calificador = 'R' or calificador = 'C' or calificador = 'O') "
                         'response.write Strsql & " strsql 4863 " & "<br>"
	                  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	                  if not Xrs.eof then
	                     XGUTABLISKEY = Xrs(0)
	                     XGUTABLISDAT = Xrs(13)

	                  END IF
                          call ofv.cerrarconsulta(Xrs)

                        ELSE

			  Strsql = "SELECT OBJSECTOR.*, ECO.descri FROM OBJSECTOR, ECO "
			  Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and eco.estado=3 and codobjeto = 0" & XGUFCOBJF
	                  Strsql = Strsql & " And (calificador = 'R' or calificador = 'C' or calificador = 'O') "
			 'response.write Strsql & " strsql 4877 " & "<br>"
			  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			  do while not Xrs.eof

			     XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
			     XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(13)

			  Xrs.movenext
			  loop
			  call ofv.cerrarconsulta(Xrs)


			END IF

'' NUEVO POR VERSION 4.6 - 25-02-2015 FIN


            CASE "5"   'PROCEDIMIENTOS

                  XGUTABLAF    = "CSCEventoSectorPerfil"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM CSCEventoSectorPerfil WHERE CODPROCED = '" & XGUFNIV2F & "' "
                    'response.write Strsql & " strsql 4902 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(2) & "."  & Xrs(6) & ".&nbsp;" & Xrs(4) & "&nbsp;-&nbsp;"  & Xrs(3) & "&nbsp;-&nbsp;" & Xrs(5)


                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM CSCEventoSectorPerfil WHERE obcodobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by SECUEN, SECUENCIA "
                    'response.write Strsql & " strsql 4918 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(2) & "."  & Xrs(6) & ".&nbsp;" & Xrs(4) & "&nbsp;-&nbsp;"  & Xrs(3) & "&nbsp;-&nbsp;" & Xrs(5)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF







            CASE "7"   'CONTROLES

                  XGUTABLAF    = "cscontrolsel"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM cscontrolsel WHERE OBJID = 0" & XGUFNIV2F
                    'response.write Strsql & " strsql 4976 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(1)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM cscontrolsel WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by DESCRI "
                    'response.write Strsql & " strsql 4991 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(1)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF





            CASE "8"   'ESQUEMAS

                  XGUTABLAF    = "csesquemasel"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM csesquemasel WHERE OBJID = 0" & XGUFNIV2F
                    'response.write Strsql & " strsql 4993 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(1)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM csesquemasel WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by DESCRI "
                    'response.write Strsql & " strsql 5008 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(1)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF






            CASE "9"   'TABLAS

                  XGUTABLAF    = "ObjTabDef"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM ObjTabDef WHERE OBJID = 0" & XGUFNIV2F
                    'response.write Strsql & " strsql 5038 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(4)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM ObjTabDef WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by tabsecuen "
                    'response.write Strsql & " strsql 5053 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(4)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF

  CASE "10" 'documentos

  CASE "13" 'contratos

         XGUTABLAF = "Objhtm"

         IF XGUFOPC = "E" THEN

            Strsql = "SELECT * FROM Objhtm WHERE OBJID = 0" & XGUFNIV2F
           'response.write Strsql & " strsql 5077 " & "<br>"
            SET Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
            IF NOT Xrs.EOF THEN
               XGUTABLISKEY = Xrs(0)
               XGUTABLISDAT = Xrs(3) & "." & Xrs(4)

            END IF
            CALL ofv.cerrarconsulta(Xrs)

         ELSE

            Strsql = "SELECT * FROM Objhtm WHERE codobjeto = 0" & XGUFCOBJF
		    Strsql = Strsql & " order by secuen "
                   'response.write Strsql & " strsql 5090 " & "<br>"
		             SET Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

		             DO WHILE NOT Xrs.EOF

		                XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
		                XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP &  Xrs(3) & "." & Xrs(4)

		                Xrs.movenext
		             LOOP
		             CALL ofv.cerrarconsulta(Xrs)


         END IF




            CASE "11"   'CONTENIDO

                  XGUTABLAF    = "OBJESTR"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM OBJESTR WHERE OBJID = 0" & XGUFNIV2F
                    'response.write Strsql & " strsql 5073 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM OBJESTR WHERE codobjeto = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by CONsecuen "
                    'response.write Strsql & " strsql 5088 " & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF



            CASE ELSE

                  XGUTABLAF    = "NO"

                 'response.write " salida por CASE ELSE strsql 5111 " & "<br>"

     END SELECT




     GETLISTNV2 = "OK"

 end function


 function GETLISTNV3(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F,XGUFNIV3F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO "

     XGUTABSEP    = CHR(9)




     IF XGUFOPC = "E" THEN

        IF  XGUFNIV1F = "9" THEN


            Strsql = "SELECT * FROM ObjTabblk WHERE OBJID = 0" & XGUFNIV3F

            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not Xrs.eof then
               XGUTABLISKEY = Xrs(0)
               XGUTABLISDAT = Xrs(5)

            END IF
            call ofv.cerrarconsulta(Xrs)


        END IF

        IF  XGUFNIV1F = "11" THEN

            Strsql = "SELECT * FROM OBJESTDET WHERE OBJID = 0" & XGUFNIV3F
            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not Xrs.eof then
               XGUTABLISKEY = Xrs(0)
               XGUTABLISDAT = UCASE(Xrs(7))
            END IF
            call ofv.cerrarconsulta(Xrs)

        END IF

     ELSE

        IF  XGUFNIV1F = "9" THEN


                     Strsql = "SELECT * FROM ObjTabblk WHERE tabdef = 0" & XGUFNIV2F
                     Strsql = Strsql & " order by tabblksec "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(5)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)

        end if


        IF  XGUFNIV1F = "11" THEN

            strsql = " SELECT * FROM CSESTDET "
            strsql = strsql & " where detcon = " & XGUFNIV2F & "  "
            strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
            strsql = strsql & " and not ( "
            strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
            strsql = strsql & " ) and stmodi = '' AND DETTIPO < 'G' "
            strsql = strsql & " order by TSEC"

            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            do while not Xrs.eof

               XGUINDN = "-&nbsp;"

               SELECT CASE XRS(5)

                      CASE "B"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;"


                      CASE "C"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


                      CASE "D"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


                      CASE "E"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"


                      CASE "F"

                            XGUINDN = "-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"



               END SELECT

               XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
               XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & XGUINDN & UCASE(Xrs(7))

            Xrs.movenext
            loop
            call ofv.cerrarconsulta(Xrs)


        END IF


     END IF

     GETLISTNV3 = "OK"

 end function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BUSCADOR DE NIVELES DE COMPONENTES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


 function GETLISTCMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF)

     XGUTABLISKEY  = "0"
     XGUTABLISKEY2 = "NO"
     XGUTABLISDAT  = "SIN DEFINIR"

     XGUTABSEP    = CHR(9)

     IF XGUFOPC = "E" THEN

        Strsql = "Select * From csacomple where CODCOMP = 0" & XGUFCOBJF
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY  = Xrs2(0)
           XGUTABLISKEY2 = Xrs2(1)
           XGUTABLISDAT  = Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)


        END IF
        call ofv.cerrarconsulta(Xrs2)



   ELSEIF XGUFOPC = "TI" THEN

        Strsql = "Select * From csacomple where tipoCOMP = '" & XGUFTOBJF & "' order by descri, TIPODESCRI"
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        do while not Xrs2.eof

           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



        Xrs2.movenext
        loop
        call ofv.cerrarconsulta(Xrs2)

     ELSEIF XGUFOPC = "TX" THEN

        Strsql = "Select * From csacomple where descri like('%" & XGUFTOBJF & "%') order by descri, TIPODESCRI"
    '''    response.write "por texto: " & strsql & " <br>"
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        do while not Xrs2.eof

           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



        Xrs2.movenext
        loop
        call ofv.cerrarconsulta(Xrs2)

     ELSEIF XGUFOPC = "TL" THEN

        Strsql = "Select * From csacomple where descri like('" & XGUFTOBJF & "%') order by descri, TIPODESCRI"
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        do while not Xrs2.eof

           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



        Xrs2.movenext
        loop
        call ofv.cerrarconsulta(Xrs2)


     ELSE

	        Strsql = "Select * From csacomple  order by descri, TIPODESCRI"
	        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	        do while not Xrs2.eof

	           XGUTABLISKEY  = XGUTABLISKEY & XGUTABSEP & Xrs2(0)
	           XGUTABLISKEY2 = XGUTABLISKEY2 & XGUTABSEP & Xrs2(1)
	           XGUTABLISDAT  = XGUTABLISDAT & XGUTABSEP & Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)



	        Xrs2.movenext
	        loop
        call ofv.cerrarconsulta(Xrs2)

     END IF

     GETLISTCMP = "OK"


 end function







 function GETLISTSLPCMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL COMPONENTE"

     XGUTABSEP    = CHR(9)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' BUSCA LAS SOLAPAS DEL COMPONENTE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  dim stabla(20)
  dim solord(20)
  dim Nombres(20)


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

  dim Solapas(20)


  for z=0 to 20
      Solapas(z) = "N"
  next


  Strsql = "Select * From tipocomp Where tipo = '" & XGUFTOBJF & "' "
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not Xrs.eof then
     sPerfil  = Xrs(4)
  end if

  call ofv.cerrarconsulta(Xrs)



  Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
'' response.write strsql & "<br>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not Xrs.eof then



    for xx = 1 to 20
        solord(xx) = 17
        Nombres(xx) = "N"
    next



    for xx = 2 to 16
        stabla (xx - 1) = Xrs(xx + 30)
    next



    yy = 0
    for xx = 1 to 15

        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           solord (stabla (xx)) = xx
        end if
    next
    tyy = yy

    xstrsol     = 1



  X=0
  for z = 2 to 16

    if Xrs(z) Then

      if  z <> 21 and z <> 4 and z <> 61 then
          x = x + 1
          Solapas(z-1)  = REPLACE(XRS(z + 15)," ","*")

      end if
    end if

  next

  end if

  call ofv.cerrarconsulta(Xrs)


  xfinal = ubound(Solapas)


     IF XGUFOPC = "E" THEN

        Strsql = "Select * From CSACOMPLE where CODCOMP = 0" & XGUFCOBJF
        set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT Xrs2.eof THEN

           XGUTABLISKEY = Xrs2(0)
           XGUTABLISDAT = Xrs2(3) & ":" & Xrs2(2) & "&nbsp;-&nbsp;" & Xrs2(6)


        END IF
        call ofv.cerrarconsulta(Xrs2)

        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then
            '   response.write solord(x)
               if clng(XGUFNIV1F) = clng(solord(x)) then

                 XGUTABLISKEY = solord(x)
                 XGUTABLISDAT = solapas(solord(x))

                 EXIT FOR

               END IF
             end if
        next



     ELSE


        for x = xstrsol to xfinal
            if solapas(solord(x)) <> "N" then

                 XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & solord(x)
                 XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP  & solord(x) & "&nbsp;" & solapas(solord(x))

            end if
        next




     END IF

     GETLISTSLPCMP = "OK"

 end function


 function GETLISTNV2CMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO " & XGUFNIV1F

     XGUTABSEP    = CHR(9)

     XGUTABLAF    = "NO"


     SELECT CASE XGUFNIV1F

            CASE "2"   'estructura

                 XGUTABLAF    = "COMPESTR"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by OBJID "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)
                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & Xrs(3)



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF







            CASE "4"   'imagenes

                  XGUTABLAF    = "COMPIMAG"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(4)

                        IF XGUTABLISDAT = "" THEN XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by SECUENCIA, OBJID "
                  ''   response.write strsql & "<br>"
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)

                        WKXGUTABLISDAT = Xrs(4)

                        IF WKXGUTABLISDAT = "" THEN WKXGUTABLISDAT = Xrs(3)

                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & WKXGUTABLISDAT



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF







            CASE "6"   'archivos

                  XGUTABLAF    = "COMPARC"

                  IF XGUFOPC = "E" THEN


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE OBJID = 0" & XGUFNIV2F

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        XGUTABLISKEY = Xrs(0)
                        XGUTABLISDAT = Xrs(4)

                        IF XGUTABLISDAT = "" THEN XGUTABLISDAT = Xrs(3)

                     END IF
                     call ofv.cerrarconsulta(Xrs)


                  ELSE


                     Strsql = "SELECT * FROM " & XGUTABLAF & " WHERE CODCOMP = 0" & XGUFCOBJF
                     Strsql = Strsql & " order by SECUENCIA, OBJID "

                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                     do while not Xrs.eof

                        XGUTABLISKEY = XGUTABLISKEY & XGUTABSEP & Xrs(0)

                        WKXGUTABLISDAT = Xrs(4)

                        IF WKXGUTABLISDAT = "" THEN WKXGUTABLISDAT = Xrs(3)

                        XGUTABLISDAT = XGUTABLISDAT & XGUTABSEP & WKXGUTABLISDAT



                    Xrs.movenext
                    loop
                    call ofv.cerrarconsulta(Xrs)


                  END IF








            CASE ELSE

                  XGUTABLAF    = "NO"



     END SELECT




     GETLISTNV2CMP = "OK"

 end function


 function GETLISTNV3CMP(XGUFOPC,XGUFCOBJF,XGUFTOBJF,XGUFNIV1F,XGUFNIV2F,XGUFNIV3F)

     XGUTABLISKEY = "0"
     XGUTABLISDAT = "TODO EL CONTENIDO "

     XGUTABSEP    = CHR(9)

     GETLISTNV3CMP = "OK"

 end function

 function GenBuscaCMP(XGUCOBJI,XGUTOBJI,XGUN1I,XGUN2I,XGUN3I,XGUN4I,XGUN5I)


 XGUFBURL = request.servervariables("SCRIPT_NAME")

 XGUFBURL = RIGHT(XGUFBURL,LEN(XGUFBURL) - (INSTRREV(XGUFBURL,"/",-1))) & "?" & REQUEST.QUERYSTRING

 XGUMASTER = REQUEST.FORM("XGUMASTER")
 XGUITEMCD = REQUEST.FORM("XGUITEMCD")
 XGUFILTRO = REQUEST.FORM("XGUFILTRO")

 XGUSTYLS = " STYLE='border-style:solid;border-width:1;border-color:#F7AA22;cursor:pointer;' "
 XGUSTYLN = " STYLE='cursor:pointer;' "


 '''IF  XGUMASTER = "TX" THEN XGUITEMCD = REQUEST.FORM("XGUBTEXTO")


 IF  XGUMASTER = "" THEN

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


     if   XGUCOBJI <> "0" THEN
          XGUFCOBJ = XGUCOBJI
          XGUFTOBJ = XGUTOBJI
     END IF

     if   XGUN1I <> "0" THEN
          XGUFNIV1 = XGUN1I
          XGUMASTER = "1"
     END IF

     if   XGUN2I <> "0" THEN
          XGUFNIV2 = XGUN2I
          XGUMASTER = "2"
     END IF

     if   XGUN3I <> "0" THEN
          XGUFNIV3 = XGUN3I
          XGUMASTER = "3"
     END IF

     if   XGUN4I <> "0" THEN
          XGUFNIV4 = XGUN4I
          XGUMASTER = "4"
     END IF

     if   XGUN5I <> "0" THEN
          XGUFNIV5 = XGUN5I
          XGUMASTER = "5"
     END IF


 ELSE

     XGUFCOBJ = REQUEST.FORM("XGUFCOBJ")
     XGUFTOBJ = REQUEST.FORM("XGUFTOBJ")
     XGUFNIV1 = REQUEST.FORM("XGUFNIV1")
     XGUFNIV2 = REQUEST.FORM("XGUFNIV2")
     XGUFNIV3 = REQUEST.FORM("XGUFNIV3")
     XGUFNIV4 = REQUEST.FORM("XGUFNIV4")
     XGUFNIV5 = REQUEST.FORM("XGUFNIV5")

    'response.write XGUFCOBJ & " XGUFCOBJ " & "<br>"
    'response.write XGUFTOBJ & " XGUFTOBJ " & "<br>"

 END IF



 XGUCONTEMS = ""



 XGUFNC1 = ""
 XGUFNC2 = ""


 set xgucn = ofv.conectar(ofv.strconn2)

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


 ' -------------------------

''XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TI','NO');document.getElementById('XGUFRM').submit();return false;"" >"
'' "clactfrm('" & XGUMASTER & "','TI','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

 '' XGUFBBOTi1 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TODOS' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','A','NO','NO','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
  XGUFBBOTi2 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TIPO' CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TI','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
  XGUFBBOTi3 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='LETRA'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TL','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"
  XGUFBBOTi4 = "<td class=ut  ><INPUT TYPE='BUTTON' VALUE='TEXTO'  CLASS=BT ONCLICK=""clactfrm('" & XGUMASTER & "','0','0','0','0','0','0','0','TX','NO');document.getElementById('XGUFRM').submit();return false;"" ></td>"

  XGUFBBOT1  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI1 & "</FORM>"
  XGUFBBOT2  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI2 & "</FORM>"
  XGUFBBOT3  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI3 & "</FORM>"
  XGUFBBOT4  = "<FORM ACTION='" & XGUFBURL & "' METHOD='POST' >" & XGUFBBOTI4 & "</FORM>"



         XGUFILDSC = "&nbsp;componentes"

  SELECT CASE XGUFILTRO

         CASE "NO"

         CASE "TI"

              XGUFILDSC1 = ":&nbsp;TIPO COMPONENTE"

              xTrColor = "#FFFFFF"

              IF XGUITEMCD <> "NO" THEN
                  XGUFNC1 = " document.getElementById('***" & XGUITEMCD & "').scrollIntoView(true); " & chr(13)
              ELSE
                  XGUFNC1 = ""
              END IF


              IF XGUCODIGO <> "NO" THEN
                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                 XGUFNC2 = ""
              END IF


              strsql = "Select * From TIPOCOMP  order by descri "
            ''response.write strsql & " strsql " & "<br>"
              set xgurs = ofv.crearconsultaEx(StrSql,xgucn,1,volver)

              do until xgurs.eof


                 XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                 XGUFBTD1 = XGUFBTD1 & " >"

                 XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UTI TITLE='Seleccionar item " & xgurs(1) & ".' ID='***" & xgurs(0) & "' "

                 IF CSTR(xgurs(0)) = XGUITEMCD THEN
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                 ELSE
                    XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                 END IF



          '''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & xgurs(0) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"
                 XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TI','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & xgurs(0) & "');document.getElementById('XGUFRM').submit();return false;"" >"


                 XGUFBTD1 = XGUFBTD1 & xgurs(1) & "</TD>"
                 XGUFBTD1 = XGUFBTD1 & "</TR>"


                          if xTrColor = "#FFFFFF" then
                             xTrColor = "#F9F9F9"
                          ELSE
                             xTrColor = "#FFFFFF"
                          END IF


              xgurs.movenext
              loop

              call ofv.cerrarconsulta(xgurs)


             XGUFBTD1xxx = XGUFBTD1






         CASE "TL"

              XGUFILDSC1 = ":&nbsp;LETRA"

              xTrColor = "#F9F9F9"


              IF XGUCODIGO <> "NO" THEN
                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                 XGUFNC2 = ""
              END IF


              dim ATABXGULTR(28)

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


              FOR XGUINX=1 TO UBOUND(ATABXGULTR) STEP 4

                  XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                  XGUFBTD1 = XGUFBTD1 & " >"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  IF ATABXGULTR(XGUINX) = XGUITEMCD THEN
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                  ELSE
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                  END IF

                  XGUFBTD1 = XGUFBTD1 & " >"

                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX) & "&nbsp;&nbsp;' CLASS=BT "
            ''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


                  XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX) & "');document.getElementById('XGUFRM').submit();return false;"" >"





                  XGUFBTD1 = XGUFBTD1 & "</TD>"


                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  IF ATABXGULTR(XGUINX+1) = XGUITEMCD THEN
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                  ELSE
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                  END IF

                  XGUFBTD1 = XGUFBTD1 & " >"
                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+1) & "&nbsp;&nbsp;' CLASS=BT "
            ''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+1) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



                  XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX+1) & "');document.getElementById('XGUFRM').submit();return false;"" >"





                  XGUFBTD1 = XGUFBTD1 & "</TD>"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  if  ATABXGULTR(XGUINX+2) = "NO" then

                     XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                  else
                     IF ATABXGULTR(XGUINX+2) = XGUITEMCD THEN
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                     ELSE
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                     END IF
                     XGUFBTD1 = XGUFBTD1 & " >"
                     XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+2) & "&nbsp;&nbsp;' CLASS=BT "
               ''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+2) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"



                  XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX+2) & "');document.getElementById('XGUFRM').submit();return false;"" >"






                  end if

                  XGUFBTD1 = XGUFBTD1 & "</TD>"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.' align=center "

                  if  ATABXGULTR(XGUINX+3) = "NO" then

                     XGUFBTD1 = XGUFBTD1 & " >&nbsp;"

                  else
                     IF ATABXGULTR(XGUINX+3) = XGUITEMCD THEN
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                     ELSE
                        XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                     END IF
                     XGUFBTD1 = XGUFBTD1 & " >"
                     XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON  VALUE='&nbsp;&nbsp;" & ATABXGULTR(XGUINX+3) & "&nbsp;&nbsp;' CLASS=BT "
              '''''''XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "','" & ATABXGULTR(XGUINX+3) & "','NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


                  XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TL','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO','" & ATABXGULTR(XGUINX+3) & "');document.getElementById('XGUFRM').submit();return false;"" >"




                 end if

                  XGUFBTD1 = XGUFBTD1 & "</TD>"



                  XGUFBTD1 = XGUFBTD1 & "</TR>"


                  if xTrColor = "#FFFFFF" then
                     xTrColor = "#F9F9F9"
                  ELSE
                     xTrColor = "#FFFFFF"
                  END IF

              NEXT









         CASE "TX"

              XGUFILDSC1 = ":&nbsp;TEXTO"



              xTrColor = "#FFFFFF"


              IF XGUCODIGO <> "NO" THEN
                 XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
              ELSE
                 XGUFNC2 = ""
              END IF



                  XGUFBTD1 = XGUFBTD1 & "<TR bgcolor=" & xTrColor & " "
                  XGUFBTD1 = XGUFBTD1 & " >"

                  XGUFBTD1 = XGUFBTD1 & "<TD CLASS=UT TITLE='SELECCIONAR FILTRO.'  "

                  IF XGUITEMCD <> "NO" THEN
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLS
                  ELSE
                     XGUFBTD1 = XGUFBTD1 & XGUSTYLN
                  END IF

                  XGUFBTD1 = XGUFBTD1 & " >"
                  XGUFBTD1 = XGUFBTD1 & "<INPUT id='XGUBTEXTO' name='XGUBTEXTO' TYPE=TEXT VALUE='" & XGUITEMCD & "' CLASS=DT size=40 >&nbsp;&nbsp;"

                  XGUFBTD1 = XGUFBTD1 & "<INPUT TYPE=BUTTON VALUE='OK' CLASS=BT "
              ''    XGUFBTD1 = XGUFBTD1 & " ONCLICK=""clactfrm('" & XGUMASTER & "','" & XGUFILTRO & "',document.getElementById('XGUBTEXTO').value,'NO','NO');document.getElementById('XGUFRM').submit();return false;"" >"


                 XGUFBTD1 = XGUFBTD1  & " ONCLICK=""clactfrm('TX','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0','NO',document.getElementById('XGUBTEXTO').value);document.getElementById('XGUFRM').submit();return false;"" >"




                  XGUFBTD1 = XGUFBTD1 & "</TD>"

                  XGUFBTD1 = XGUFBTD1 & "</TR>"





         CASE ELSE

              XGUFILDSC1 = "&nbsp;"

              XGUFILTRO = "NO"
              XGUITEMCD = "NO"
              XGUCODIGO = "NO"
              XGUCODDSC = "NO"


  END SELECT






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

  SELECT CASE XGUMASTER


           '  IF XGUCODIGO <> "NO" THEN
           '     XGUFNC2 = " document.getElementById('---" & XGUCODIGO & "').scrollIntoView(true); " & chr(13)
           '  ELSE
           '     XGUFNC2 = ""
           '  END IF

       CASE "TI","TL","TX"




	                   CALL GETLISTCMP(XGUMASTER,"0",XGUITEMCD)


	                   XGUTABKEY  = SPLIT(XGUTABLISKEY,CHR(9))
	                   XGUTABKEY2 = SPLIT(XGUTABLISKEY2,CHR(9))
	                   XGUTABDAT  = SPLIT(XGUTABLISDAT,CHR(9))

	                   xTrColor = "#FFFFFF"



	                   IF XGUFCOBJ <> "0" THEN
	                      XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & chr(13)
	                   ELSE
	                      XGUFNC2 = ""
	                   END IF



	                   FOR XGUINX=0 TO UBOUND(XGUTABKEY)


	                       XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
	                       XGUFBTD2 = XGUFBTD2 & " >"

	                       XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR COMPONENTE.'  ID='---" & XGUTABKEY(XGUINX) & "' "

	                       IF CSTR(XGUTABKEY(XGUINX)) = cstr(XGUFCOBJ) THEN
	                          XGUFBTD2 = XGUFBTD2 & XGUSTYLS
	                       ELSE
	                          XGUFBTD2 = XGUFBTD2 & XGUSTYLN
	                       END IF



	                       XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
	                       XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
	                       XGUFBTD2 = XGUFBTD2 & "</TR>"


	                            if xTrColor = "#FFFFFF" then
	                               xTrColor = "#F9F9F9"
	                            ELSE
	                               xTrColor = "#FFFFFF"
	                            END IF


	                   NEXT





        CASE "0"

                CALL GETLISTCMP("L","0","0")
                XGUTABKEY  = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABKEY2 = SPLIT(XGUTABLISKEY2,CHR(9))
                XGUTABDAT  = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"



                IF XGUFCOBJ <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFCOBJ & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR COMPONENTE.'  ID='---" & XGUTABKEY(XGUINX) & "' "

                    IF CSTR(XGUTABKEY(XGUINX)) = cstr(XGUFCOBJ) THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF



                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('1','" & XGUTABKEY(XGUINX) & "','" & XGUTABKEY2(XGUINX) & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT




        CASE "1"


                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("L",XGUFCOBJ,XGUFTOBJ,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"


                IF XGUFNIV1 <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFNIV1 & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR SOLAPA.'  ID='---" & XGUTABKEY(XGUINX) & "' "


                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV1 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF

                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUTABKEY(XGUINX) & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT



        CASE "2"

                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"





                CALL GETLISTNV2CMP("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"


                IF XGUFNIV2 <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFNIV2 & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF


                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 1ER. NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV2 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('3','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUTABKEY(XGUINX) & "','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT





        CASE "3"

                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"





                CALL GETLISTNV3CMP("L",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,"0")
                XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9))
                XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9))

                xTrColor = "#FFFFFF"

                IF XGUFNIV3 <> "0" THEN
                   XGUFNC2 = " document.getElementById('---" & XGUFNIV3 & "').scrollIntoView(true); " & chr(13)
                ELSE
                   XGUFNC2 = ""
                END IF



                FOR XGUINX=0 TO UBOUND(XGUTABKEY)


                    XGUFBTD2 = XGUFBTD2 & "<TR bgcolor=" & xTrColor & " "
                    XGUFBTD2 = XGUFBTD2 & " >"

                    XGUFBTD2 = XGUFBTD2 & "<TD CLASS=UTI TITLE='SELECCIONAR 2DO.NIVEL '  ID='---" & XGUTABKEY(XGUINX) & "' "



                    IF CSTR(XGUTABKEY(XGUINX)) = XGUFNIV3 THEN
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLS
                    ELSE
                       XGUFBTD2 = XGUFBTD2 & XGUSTYLN
                    END IF


                    XGUFBTD2 = XGUFBTD2 & " ONCLICK=""clactfrm('4','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','" & XGUTABKEY(XGUINX) & "','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                    XGUFBTD2 = XGUFBTD2 & XGUTABDAT(XGUINX) & "</TD>"
                    XGUFBTD2 = XGUFBTD2 & "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF


                NEXT


        CASE "4"

                CALL GETLISTCMP("E",XGUFCOBJ,XGUFTOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT  "
                XGUFBTDF0 = XGUFBTDF0 & " ONCLICK=""clactfrm('0','" & XGUFCOBJ & "','" & XGUFTOBJ & "','0','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF0 = XGUFBTDF0 & XGUImG & "&nbsp;"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"


                CALL GETLISTSLPCMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                XGUFBTDF1 = XGUFBTDF1 & " >"

                XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT  "
                XGUFBTDF1 = XGUFBTDF1 & " ONCLICK=""clactfrm('1','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','0','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF1 = XGUFBTDF1 & XGUImG & "&nbsp;"
                XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF1 = XGUFBTDF1 & "</TR>"



                CALL GETLISTNV2CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                XGUFBTDF2 = XGUFBTDF2 & " >"

                XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT  "
                XGUFBTDF2 = XGUFBTDF2 & " ONCLICK=""clactfrm('2','" & XGUFCOBJ & "','" & XGUFTOBJ & "','" & XGUFNIV1 & "','" & XGUFNIV2 & "','0','0','0');document.getElementById('XGUFRM').submit();return false;"" >"
                XGUFBTDF2 = XGUFBTDF2 & XGUImG & "&nbsp;"
                XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF2 = XGUFBTDF2 & "</TR>"



                CALL GETLISTNV3CMP("E",XGUFCOBJ,XGUFTOBJ,XGUFNIV1,XGUFNIV2,XGUFNIV3)
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


 END SELECT



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'XGUMASTER

 IF XGUMASTER > "0" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


 END IF

 IF XGUMASTER > "1" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


 END IF

 IF XGUMASTER > "2" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


 END IF

 IF XGUMASTER > "3" THEN

    XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
    XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


 END IF

 IF XGUMASTER > "0" THEN

    XGUFBDIV1  = "<DIV id='BSUSR4' STYLE='margin-top:1;margin-left:1;width:99%;background-color:#ffffff;"
    XGUFBDIV1  =  XGUFBDIV1  & "border-style:solid;border-width:1;border-color:#eeeeee;'>"
    XGUFBDIV1  =  XGUFBDIV1  & XGUFBTAB1 & "</DIV>"

 END IF




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


  call ofv.cerrarconn(xgucn)


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
  Whtml = Whtml & XGUFBBOT1
  Whtml = Whtml & XGUFBBOT2
  Whtml = Whtml & XGUFBBOT3
  Whtml = Whtml & XGUFBBOT4
  Whtml = Whtml & XGUFBBOT5

  if XGUCODDSC <> "NO" then
     Whtml = Whtml & "<td class=uti width=70% style='border-style:solid;border-width:1;border-color:#cccccc;' >" & XGUCODDSC & "</td>"
  else
     Whtml = Whtml & "<td class=uti width=70%  >&nbsp;</td>"
  end if

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
 Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4,XGU5,XGU6,XGU7,XGU8,XGU9) { " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUMASTER').value=XGU0; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFCOBJ').value=XGU1; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFTOBJ').value=XGU2; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV1').value=XGU3; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV2').value=XGU4; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV3').value=XGU5; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV4').value=XGU6; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFNIV5').value=XGU7; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUFILTRO').value=XGU8; " & chr(13)
 Whtml = Whtml & "    document.getElementById('XGUITEMCD').value=XGU9; " & chr(13)

 Whtml = Whtml & " } " & chr(13)
 Whtml = Whtml & XGUFNC1
 Whtml = Whtml & XGUFNC2
 Whtml = Whtml & "</script>"







     GenBuscaCMP = Whtml


 end function






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function SETEMPRE()

     xGRP = "S"
     xUSR = "S"

  set xcn = ofv.conectar(ofv.strconn2)

  Strsql = "select SITE from USUARIOS where ID = '" & SESSION("USUARIO") & "'"
  set Xrs = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if not Xrs.eof then
     USREMP = Xrs(0)
     USREMP = clng(USREMP)
  end if

  Strsql = "select * from empreatrm where ATRMTIPO = 'G' AND atrmusrgrp = '" & SESSION("SEGGRP") & "' ORDER BY EMPRE "
  set Xrs1 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if Xrs1.eof then
     xGRP = "N"
  end if


  Strsql = "select * from empreatrm where ATRMTIPO = 'U' AND atrmusrgrp = '" & SESSION("USUARIO") & "' ORDER BY EMPRE "
  set Xrs2 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if Xrs2.eof then
     xUSR = "N"
  end if

        ATREMP(0,0) = 0

  Strsql = "select * from CLIENTEPLAN ORDER BY ID"
  set Xrs = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if not Xrs.eof then
     DO WHILE NOT XRS.EOF
        IAEMP = IAEMP + 1
        TAEMP = IAEMP

        ATREMP(IAEMP,0) = XRS(0)
        ATREMP(IAEMP,1) = XRS(6)
        ATREMP(IAEMP,4) = XRS(1)
        ATREMP(IAEMP,5) = XRS(5)
        ATREMP(IAEMP,2) = 1
        ATREMP(IAEMP,3) = 1

       FOR XX1=0 TO 50
        IF XGRP <> "N" THEN
           IF  XRS1(0) = XRS(0) THEN
               ATREMP(IAEMP,2) = ABS(CINT(XRS1(5)))
               ATREMP(IAEMP,3) = ABS(CINT(XRS1(6)))
               XRS1.MOVENEXT
               IF XRS1.EOF THEN XGRP = "N"
               EXIT FOR
           ELSEIF  XRS1(0) < XRS(0) THEN
               XRS1.MOVENEXT
               IF XRS1.EOF THEN XGRP = "N"
           END IF
        ELSE
           EXIT FOR
        END IF
       NEXT

       FOR XX1=0 TO 50
        IF XUSR <> "N" THEN
           IF  XRS2(0) = XRS(0) THEN
               ATREMP(IAEMP,2) = ABS(CINT(XRS2(5)))
               ATREMP(IAEMP,3) = ABS(CINT(XRS2(6)))
               XRS2.MOVENEXT
               IF XRS2.EOF THEN XUSR = "N"
               EXIT FOR
           ELSEIF  XRS2(0) < XRS(0) THEN
               XRS2.MOVENEXT
               IF XRS2.EOF THEN XUSR = "N"
           END IF
        ELSE
           EXIT FOR
        END IF
       NEXT


       IF XRS(0) = USREMP THEN
          IF  ATREMP(IAEMP,2) = 1 THEN

              ATREMP(0,0) = ATREMP(IAEMP,0)
              ATREMP(0,1) = ATREMP(IAEMP,1)
              ATREMP(0,4) = ATREMP(IAEMP,4)
              ATREMP(0,5) = ATREMP(IAEMP,5)
              ATREMP(0,2) = ATREMP(IAEMP,2)
              ATREMP(0,3) = ATREMP(IAEMP,3)

          END IF
       END IF


     XRS.MOVENEXT
     LOOP

     IF ATREMP(0,0) = 0 THEN

        FOR IAEMP = 1 TO TAEMP
          IF  ATREMP(IAEMP,2) = 1 THEN

              ATREMP(0,0) = ATREMP(IAEMP,0)
              ATREMP(0,1) = ATREMP(IAEMP,1)
              ATREMP(0,4) = ATREMP(IAEMP,4)
              ATREMP(0,5) = ATREMP(IAEMP,5)
              ATREMP(0,2) = ATREMP(IAEMP,2)
              ATREMP(0,3) = ATREMP(IAEMP,3)
              EXIT FOR
          END IF


        NEXT
     END IF

    DIM WATREMPEL(5)
    DIM WATREMPTB(50)


        FOR IAEMP = 0 TO TAEMP
         FOR IAEMP2 = 0 TO 5
              WATREMPEL(IAEMP2) = cstr(ATREMP(IAEMP,IAEMP2) &"")
         NEXT

         WATREMPTB(IAEMP) = JOIN(WATREMPEL,CHR(09))

        NEXT


    SESSION("ATREMP") = JOIN(WATREMPTB,CHR(10))
    SESSION("TAEMP")  = TAEMP
    SESSION("SETEMP") = "S"


  end if



end function

FUNCTION NEWEMPRE(XEMPRE)

    IF XEMPRE = "NO" THEN

       if SESSION("SETEMP") <> "S" THEN   xxxx = SETEMPRE()

       WATREMPTB = SPLIT(SESSION("ATREMP"),CHR(10))

       FOR IAEMP=0 TO UBOUND(WATREMPTB)

           WATREMPEL = SPLIT(WATREMPTB(IAEMP),CHR(09))

           FOR IAEMP2=0 TO UBOUND(WATREMPEL)

               ATREMP(IAEMP,IAEMP2) = WATREMPEL(IAEMP2)

           NEXT

       NEXT

       TAEMP  = SESSION("TAEMP")

       NEWEMPRE = "OK"

    ELSE

     NEWEMPRE = "N"

     IF ATREMP(0,0) <> XEMPRE THEN

        NEWEMPRE = "S"

        FOR IAEMP = 1 TO TAEMP
          IF  ATREMP(IAEMP,0) = XEMPRE THEN

              ATREMP(0,0) = ATREMP(IAEMP,0)
              ATREMP(0,1) = ATREMP(IAEMP,1)
              ATREMP(0,4) = ATREMP(IAEMP,4)
              ATREMP(0,5) = ATREMP(IAEMP,5)
              ATREMP(0,2) = ATREMP(IAEMP,2)
              ATREMP(0,3) = ATREMP(IAEMP,3)
              EXIT FOR
          END IF


        NEXT

        DIM WATREMPEL(5)
        DIM WATREMPTB(50)


        FOR IAEMP = 0 TO TAEMP
         FOR IAEMP2 = 0 TO 5
              WATREMPEL(IAEMP2) = ATREMP(IAEMP,IAEMP2)
         NEXT

         WATREMPTB(IAEMP) = JOIN(WATREMPEL,CHR(09))

        NEXT


        SESSION("ATREMP") = JOIN(WATREMPTB,CHR(10))
        SESSION("TAEMP")  = TAEMP

     END IF
    END IF

END FUNCTION


function GETGLBHMSW(XITIME)

 if not isnumeric(XITIME & "") then

 GETGLBHMSW = "000005"

 else

XITIME=clng("0" & replace(cstr(XITIME & ""),"-",""))
GLBbb = XITIME MOD 3600
GLBbC = (GLBBB) MOD 60
GLBbD = (GLBBC) MOD 1

GLBSH = (XITIME - GLBBB) / 3600
GLBSM = (GLBBB - GLBBC) / 60
GLBSS = GLBBC / 1

GLBSH = cstr(GLBSH)
GLBSM = cstr(GLBSM)
GLBSS = cstr(GLBSS)

if len(GLBSH) = 1 then GLBSH = "0" & GLBSH
if len(GLBSM) = 1 then GLBSM = "0" & GLBSM
if len(GLBSS) = 1 then GLBSS = "0" & GLBSS

 XOSH = GLBSH & ":"
 XOSM = GLBSM & ":"
 XOSS = GLBSS & ""


 GETGLBHMSW = XOSH & XOSM & XOSS

 end if

END FUNCTION


function GETGLBHMS(XITIME)

 if not isnumeric(XITIME & "") then

 GETGLBHMS = "000005"

 else
''response.write "-" & XITIME & "-" & "<br>"
XITIME=clng("0" & replace(cstr(XITIME & ""),"-",""))
GLBbb = XITIME MOD 3600
GLBbC = (GLBBB) MOD 60
GLBbD = (GLBBC) MOD 1

GLBSH = (XITIME - GLBBB) / 3600
GLBSM = (GLBBB - GLBBC) / 60
GLBSS = GLBBC / 1

IF GLBSH > 0 THEN XOSH = GLBSH & "h&nbsp;"
IF GLBSM > 0 THEN XOSM = GLBSM & "m&nbsp;"
IF GLBSS > 0 THEN XOSS = GLBSS & "s&nbsp;"


 GETGLBHMS = XOSH & XOSM & XOSS

 end if

END FUNCTION

function GETFECH()



    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

   XLHORDAT = (CINT(sh)*3600) + (CINT(sm)*60) + CINT(ss)
 ''   XLHORDAT = TIME()
  ''  RESPONSE.WRITE XLHORDAT & "<BR>"
    XLFECDAT = CLNG(DATE())



END FUNCTION

function GETNWEN(XNCODIGO,XNTIPO)


    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

XCodobJ = sh & sm & ss

GThms   = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses   = session.sessionid
GTsecu  = cstr(XCodobJ & GThms & GTses)

GTsecuf = XNCODIGO & XNTIPO


for GTi=1 to len(GTsecu)
    GTcs = mid(GTsecu,GTi,1)
    if GTcs => "0" and GTcs <= "9" then
       GTsecuf = GTsecuf & GTcs
    end if
next

GETNWEN = trim(GTsecuf)

end function

function GETLOGN(XLTYPE)

    IF XLTYPE = "" THEN XLTYPE = "L"

    sh = datepart("h",TIME())
    sm = datepart("n",TIME())
    ss = datepart("s",TIME())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

SELECT CASE XLTYPE
       CASE "L"
             GTsecuf = "LOG"
       CASE "LE"
             GTsecuf = "LXT"
       CASE "LD"
             GTsecuf = "LDE"
END SELECT

for GTi=1 to len(GTsecu)
    GTcs = mid(GTsecu,GTi,1)
    if GTcs => "0" and GTcs <= "9" then
       GTsecuf = GTsecuf & GTcs
    end if
next

GETLOGN = trim(GTsecuf)

END FUNCTION


'''' LOG DE USUARIOS ''''''''''''''''''''''''''''''''''''''''''''''''''

FUNCTION GENUSRLOG()

   '      Session("Usuario") = UCase(Request.Form("Usuarnt"))
   '      Session("UsuarNT") = UCase(Request.Form("Usuarnt"))
   '      Session("DOMINIO") = UCase(Request.Form("DOMAINnt"))
   '      Session("TRUSTEDUSR") = "S"


  StrSql = "Select * From OBJALOG where CODOBJETO = 0 AND  TIPOBJETO = 'SKRLOG' "
  STRSQL = STRSQL & " AND USUARIO = '" & session("USUARIO") & "' AND  NETUSUARIO = '" & session("USUARNT") & "' "
  STRSQL = STRSQL & " AND DOMINIO = '" & session("DOMINIO") & "' "
 ''       response.write strsql & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if Xrs.eof then

        if session("APLMDCLI")="BC" THEN
		      STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED,USRNAME) VALUES ("
		      STRSQL = STRSQL & "'0','SKRLOG','NO','0','0','0','0','" & session("USUARIO") & "','" & session("USUARNT") & "','" & session("DOMINIO") & "','" & Session("TRUSTEDUSR") & "','"&Session("SegUsuario")&"')"
		 ELSE
		      STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED) VALUES ("
		      STRSQL = STRSQL & "'0','SKRLOG','NO','0','0','0','0','" & session("USUARIO") & "','" & session("USUARNT") & "','" & session("DOMINIO") & "','" & Session("TRUSTEDUSR") & "')"
         END IF
        CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)
'        response.write strsql

  END IF



END FUNCTION

FUNCTION GETUSRLOGUPD(XLTIPO)

 '' RESPONSE.WRITE XLTIPO & "<BR>"
  XLOPC = "N"

  SELECT CASE XLTIPO

         CASE "LOGON"
    ''  RESPONSE.WRITE XLTIPO & "0<BR>"
              XLOPC = "S"

              XLLOGIDDET = GETLOGN("L")
              SESSION("LOGIDUSRDET") = XLLOGIDDET
              XXXX = GETFECH()
              XXXX = GENUSRLOGDET(XLTIPO)
              XLUPD = " SET AEDI = AEDI + 1, FULTACC = '" & XLFECDAT & "' "



         CASE "LOGOFF"
              XLOPC = "S"
              XXXX = GETFECH()
              XLHORAINI = GETUSRLOGDETTIME()
              XLUPD = " SET TTOTALACC = TTOTALACC + " & (cdbl("0" & XLHORDAT) - cdbl("0" & XLHORAINI))
              XXXX = GENUSRLOGDET(XLTIPO)


  END SELECT
     '' RESPONSE.WRITE SESSION("LOGIDUSRDET") & "1<BR>"
  IF XLOPC = "S" THEN
     '' RESPONSE.WRITE SESSION("LOGIDUSRDET") & "2<BR>"
     if session("APLMDCLI")="BC" THEN
	   StrSql = "UPDATE OBJALOG " & XLUPD &",usrname='" &Session("SegUsuario")& "'"
	  ELSE
	   StrSql = "UPDATE OBJALOG " & XLUPD
     END IF

     StrSql = StrSql & " where CODOBJETO = 0 AND  TIPOBJETO = 'SKRLOG' "
     STRSQL = STRSQL & " AND USUARIO = '" & session("USUARIO") & "' AND  NETUSUARIO = '" & session("USUARNT") & "' "
     STRSQL = STRSQL & " AND DOMINIO = '" & session("DOMINIO") & "' "
    ''    response.write strsql & "<BR>"
     CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)

     'response.write StrSql



  END IF



END FUNCTION

FUNCTION GENUSRLOGDET(XLTIPO)

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
  XLDDIRI = Ucase(Request.ServerVariables("REMOTE_ADDR"))
  XLDLOGI = SESSION("LOGIDUSRDET")
  XLDDOMI = Session("DOMINIO")
  XLDTRUS = Session("TRUSTEDUSR")

  SELECT CASE XLTIPO

         CASE "LOGON"
              XLOPC = "S"
              XLDTACC = "LOGON"
              XLDFINI = XLFECDAT
              XLDHINI = XLHORDAT
              XLDFFIN = XLFECDAT
              XLDHFIN = XLHORDAT


         CASE "LOGOFF"



           '   StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
            '  StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
           ''  response.write strsql   & "<BR>"
           '   CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)





  END SELECT



  IF XLOPC = "S" THEN



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
        CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)




  END IF


              StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
              StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
           ''  response.write strsql   & "<BR>"
              CALL ofv.crearconsultaEx(StrSql,cn2,1,parametros)



END FUNCTION

FUNCTION GETUSRLOGDETTIME()

  GETUSRLOGDETTIME = XLHORDAT




  StrSql = "Select HINICIO From OBJALOGDET where LOGID = '" & SESSION("LOGIDUSRDET") & "' "
          '   response.write strsql   & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if NOT Xrs.eof then

     GETUSRLOGDETTIME = XRS(0)

  END IF



END FUNCTION

'''' LOG DE OBJETOS ''''''''''''''''''''''''''''''''''''''''''''''''''


FUNCTION GENOBJALOG()

  StrSql = "Select * From OBJALOG where CODOBJETO = " & session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
 ''       response.write strsql & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Xrs.eof then

        STRSQL = "INSERT INTO OBJALOG (CODOBJETO,TIPOBJETO,FULTACC,AEDI,ACON,CACT,TTOTALACC,USUARIO,NETUSUARIO,DOMINIO,TRUSTED) VALUES ("
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TCliente") & "','NO','0','0','0','0','NO','NO','NO','NO')"

        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

  END IF



END FUNCTION

FUNCTION GETOBJALOGUPD(XLTIPO)

 '' RESPONSE.WRITE XLTIPO & "<BR>"
  XLOPC = "N"

  SELECT CASE XLTIPO

         CASE "ED"
    ''  RESPONSE.WRITE XLTIPO & "0<BR>"
              XLOPC = "S"

              XLLOGIDDET = GETLOGN("LD")
              SESSION("LOGIDDET") = XLLOGIDDET
              XXXX = GETFECH()
              XXXX = GENLOGDET(XLTIPO)
              XLUPD = " SET AEDI = AEDI + 1, FULTACC = '" & XLFECDAT & "' "

         CASE "CON","ACCMOD"


              IF XLTIPO = "CON" THEN XLOPC = "S"

              XLLOGIDDET = GETLOGN("LD")
              SESSION("LOGIDDET") = XLLOGIDDET
              XXXX = GETFECH()
              XXXX = GENLOGDET(XLTIPO)
              XLUPD = " SET ACON = ACON + 1, FULTACC = '" & XLFECDAT & "' "

         CASE "ACT"
              XLOPC = "S"
              XLUPD = " SET CACT = CACT + 1 "

         CASE "FIN"
              XLOPC = "S"
              XXXX = GETFECH()
              XLHORAINI = GETLOGDETTIME()
              XLUPD = " SET TTOTALACC = TTOTALACC + " & (clng(XLHORDAT) - clng(XLHORAINI))
              XXXX = GENLOGDET(XLTIPO)


  END SELECT
     '' RESPONSE.WRITE SESSION("LOGIDDET") & "1<BR>"
  IF XLOPC = "S" THEN
     '' RESPONSE.WRITE SESSION("LOGIDDET") & "2<BR>"

     StrSql = "UPDATE OBJALOG " & XLUPD
     StrSql = StrSql & " where CODOBJETO = " & session("NCliente") & " AND  TIPOBJETO = '" & session("TCliente") & "' "
    ''    response.write strsql & "<BR>"
     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)



  END IF



END FUNCTION



FUNCTION GENLOGDET(XLTIPO)

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
  XLDDIRI = Ucase(Request.ServerVariables("REMOTE_ADDR"))
  XLDLOGI = SESSION("LOGIDDET")
  XLDDOMI = Session("DOMINIO")
  XLDTRUS = Session("TRUSTEDUSR")
  XLDULOG = SESSION("LOGIDUSRDET")

  SELECT CASE XLTIPO

         CASE "ED"
              XLOPC = "S"
              XLDTACC = "EDEN"
              XLDFINI = XLFECDAT
              XLDHINI = XLHORDAT
              XLDFFIN = XLFECDAT
              XLDHFIN = XLHORDAT

         CASE "CON"
              XLOPC = "S"
              XLDTACC = "CONSEN"
              XLDFINI = XLFECDAT
              XLDHINI = XLHORDAT

         CASE "ACCMOD"
              XLOPC = "S"
              XLDTACC = "ACCMOD"
              XLDFINI = XLFECDAT2
              XLDHINI = XLHORDAT2
              XLDFFIN = XLFECDAT
              XLDHFIN = XLHORDAT

         CASE "FIN"



     '         StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
     '         StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDDET") & "' "
           ''  response.write strsql   & "<BR>"
     '         CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)





  END SELECT



  IF XLOPC = "S" THEN



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
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)




  END IF


              StrSql = "UPDATE OBJALOGDET SET FFINAL = '" & XLFECDAT & "', HFINAL = " & XLHORDAT & "  "
              StrSql = StrSql & " where LOGID = '" & SESSION("LOGIDDET") & "' "
           ''  response.write strsql   & "<BR>"
              CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)


END FUNCTION

FUNCTION GETLOGDETTIME()

  GETLOGDETTIME = XLHORDAT




  StrSql = "Select HINICIO From OBJALOGDET where LOGID = '" & SESSION("LOGIDDET") & "' "
          '   response.write strsql   & "<BR>"
  set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT Xrs.eof then

     GETLOGDETTIME = XRS(0)

  END IF



END FUNCTION

FUNCTION GENLOGEXT(XLTIPO)

  XLOPC = "N"

  XLDTACC = "NO"
  XXXX    = GETFECH()
  XLDFINI = XLFECDAT
  XLDHINI = XLHORDAT
  XLDLOGE = GETLOGN("LE")
  XLDLOGI = SESSION("LOGIDDET")
  XLDULOG = SESSION("LOGIDUSRDET")

  SELECT CASE XLTIPO

         CASE "EDIDET"
              XLOPC   = "S"
              XLDTACC = XLTIPO
              XLDMOBJ = Session("Form2")
              XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
              XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

              IF (XLDCOBJ & "") = "" THEN XLOPC = "N"

             '' RESPONSE.WRITE REQUEST.QUERYSTRING & "<BR>"

         CASE "ACCDET"
              XLOPC   = "S"
              XLDTACC = XLTIPO
              XLDMOBJ = Session("Form2")
              XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
              XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")


  END SELECT



  IF XLOPC = "S" THEN


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

  set cn = ofv.conectar(ofv.strconn2)
          ''  response.write strsql   & "<BR>"
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

  call ofv.cerrarconn(cn)


  END IF





END FUNCTION

Function SqlExLog(XLTIPO,sql,CNEX1,TABLA,wkeyname,VALOR,TIPOPARM)

 IF wkeyname="" THEN
  StrsqlZX = "Select * From " & Tabla
  set rsKEY = ofv.crearconsultaEx(StrSqlZX,CNEX1,1,volver)
  wkeyname = rsKEY(0).name
  CALL OFV.CERRARCONSULTA(rsKEY	)
 END IF

STRSQLS= "PRC_EXECSQL '" & REPLACE(SQL,"'","''") & "'"

 SET VV=ofv.crearconsultaEx(StrSqlS,CNEX1,1,volver)
	if cn.errors.count = 0 then
	   IF NOT ISNULL(VV(0)) THEN
		  IDINSERT=VV(0)
		ELSE
		  IDINSERT=VALOR
	   END IF
	  xxxr=GENLOGEXTDATA(XLTIPO,TABLA,wkeyname,IDINSERT,TIPOPARM)
	end if
CALL OFV.CERRARCONSULTA(VV)
end function

FUNCTION GENLOGEXTDATA(XLTIPO,TABLA,NOMCOL,IDCOL,TIPOPARAM)

   XLOPC = "N"
   TABLA=ucase(tabla &"")
   NOMCOL=UCASE(NOMCOL &"")
   XLDTACC = "NO"
   XXXX = GETFECH()
   XLDFINI = XLFECDAT - clng(cdate("01-01-1900"))
   XLDHINI = XLHORDAT
   XLDLOGE = GETLOGN("LE")
   XLDLOGI = SESSION("LOGIDDET")
   XLDULOG = SESSION("LOGIDUSRDET")
   TIPOPARAM=UCASE(TIPOPARAM&"")
   SELECT CASE XLTIPO

      CASE "INSERT"
         XLOPC = "S"
         IF TIPOPARAM <> "" THEN
          XLDTACC = TIPOPARAM & "-" & XLTIPO
         ELSE
          XLDTACC=XLTIPO
         END IF
         XLDMOBJ = Session("CorApli")
         XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
         XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

         'IF (XLDCOBJ & "") = "" THEN XLOPC = "N"

         '' RESPONSE.WRITE REQUEST.QUERYSTRING & "<BR>"

      CASE "UPDATE"
         XLOPC = "S"
          IF TIPOPARAM <> "" THEN
		      XLDTACC = TIPOPARAM & "-" & XLTIPO
		  ELSE
		      XLDTACC=XLTIPO
         END IF
         XLDMOBJ = Session("CorApli")
         XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
         XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

     CASE "DELETE"
		  XLOPC = "S"
		 IF TIPOPARAM <> "" THEN
		   XLDTACC = TIPOPARAM & "-" & XLTIPO
		 ELSE
		    XLDTACC=XLTIPO
         END IF
		  XLDMOBJ = Session("CorApli")
		  XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
          XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

    CASE else
		  XLOPC = "S"

		   IF TIPOPARAM <> "" THEN
			 XLDTACC = TIPOPARAM & "-" & XLTIPO
		   ELSE
			 XLDTACC=XLTIPO
           END IF

		  XLDMOBJ = Session("CorApli")
		  XLDCOBJ = REQUEST.QUERYSTRING("GLBSOLN")
          XLDOOBJ = REQUEST.SERVERVARIABLES("SCRIPT_NAME")

   END SELECT
   XLDCOBJ=IDCOL
   IF XLOPC = "S" THEN

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

      CALL ofv.crearconsultaEx(STRSQLX, cn, 1, parametros)

     ' CALL ofv.cerrarconn(cnER)

   END IF
    ' SESSION("STRSQLX")= STRSQLX   &"---" & XLTIPO & "<BR>"

END FUNCTION


Function FechadorCor(wkfecha,param1)


     'if wkfecha & "" <> "" and wkfecha <> "&nbsp;" and wkfecha <> "..." then

 SELECT CASE SESSION("TIPOFECHA")


  CASE 0
     if isdate(wkfecha) then

       fec = cdate(wkfecha)
       ndia = day(fec)
       nMes = Month(fec)
       anio=year(fec)

       if ndia < 10 then ndia="0" & ndia
       if nMes < 10 then nMes="0" & nMes

       FechadorCor = ndia &"/" & nmes & "/" & anio
    else
       FechadorCor=wkfecha
    end if

   Case Else
     FechadorCor=wkfecha

  END SELECT

End Function


function  getuserDb()

    Dim oContext
    Dim myUser




    Dirc=  Request.ServerVariables("PATH_TRANSLATED")
    Dirc=left(Dirc,instrRev(Dirc,"\",-1))

    Dirc=Dirc & "incfile\Application.xml"

    'Dirc="C:\contextdelivery\Application.xml"

   ' response.write dirc
    Set oContext = Server.createobject("ContextDeliveryCOM.contextdelivery")
    set myUser	 =server.createobject("ContextDeliveryCOM.User")



    oContext.SetApplicationFile cstr(Dirc)

    Set myUser = OContext.GetUser("AliasAplicacion", "SQLALIAS", "DESA")
   ' response.write "Usuario=" & myUser.GetUserName() & "<br>"
   ' response.write "Password=" & myUser.GetPassword() & "<br>

    ofv.ADbUid=myUser.GetUserName()

    ofv.ADbPwd =ofv.getCript("E",myUser.GetPassword()&"")

    set oContext = nothing
    set myUser	 = nothing

end function



Function GTESTADODSC(GTESTADCOD)

      IF CSTR(GTESTADCOD & "") = "0" OR CSTR(GTESTADCOD & "") = "" THEN

         GTESTADODSC  = "-----"

      ELSE

         Set cnXX = OFV.conectar(OFV.strconn2)

           XXstrsql = "SELECT ESTADO FROM ESTADOS    "
           XXstrsql = XXstrsql & " WHERE OBJID = 0" & GTESTADCOD
       '''   response.write XXstrsql & "<br>"
           set XXRS = ofv.crearconsultaEx(XXStrSql,cnXX,1,parametros)
           if NOT XXRS.eof then
              XXWKESTADO = XXRS(0)
            ELSE
              XXWKESTADO = GTESTADCOD
           END IF

         GTESTADODSC  = XXWKESTADO

      END IF


End Function



Function GTFechador(GTWfecha)

      IF CSTR(GTWfecha & "") = "0" OR CSTR(GTWfecha & "") = "" THEN

         GTFechador  = "-----"

      ELSE



         GTWfec = cdate(GTWfecha + clng(cdate("01-01-1900")))
         GTWndia = ofv.cdia(Weekday(GTWfec))
         GTWnMes = ofv.cmes(Month(GTWfec))

         GTwFechas = "  " & GTWndia & ", " & Day(GTWfec) & " de " & GTWnMes & " de " & Year(GTWfec)

         GTFechador  = ucase(GTwFechas)

      END IF


End Function

Function GTFechadorsh(GTWfecha)

      IF CSTR(GTWfecha & "") = "0" OR CSTR(GTWfecha & "") = "" THEN

         GTFechadorSH  = "-----"

      ELSE


         GTwfec = cdate(GTWfecha + clng(cdate("01-01-1900")))

         GTWFDIAB = CSTR(Day(GTwfec))
         GTWFMESB = CSTR(Month(GTwfec))

         IF LEN(GTWFDIAB) = 1 THEN GTWFDIAB = "0" & GTWFDIAB
         IF LEN(GTWFMESB) = 1 THEN GTWFMESB = "0" & GTWFMESB

         GTwFechash = GTWFDIAB & "/" & GTWFMESB & "/" & Year(GTwfec)
         GTFechadorSH  = ucase(GTwFechash)

      END IF

End Function

function GTCDSPHCALR(GTCDATEW,GTCDATEWL,GTCDATEWS)

  GTWhtm = ""

  GTWhtm = GTWhtm & " onclick="" " & chr(13)
 ' GTWhtm = GTWhtm & "   alert('" & GTCDATEW & "');" & chr(13)
  GTWhtm = GTWhtm & " DXCALDATE('" & GTCDATEW & "','" & GTCDATEWL & "','" & GTCDATEWS & "') ;" & chr(13)
  GTWhtm = GTWhtm & """ " & chr(13)

  GTCDSPHCALR = GTWhtm

end function

function GTCSETCALR(GTCFRM,GTCDAT,GTCSMT,GTCFCH,GTCFCHDAT,GTCIN,GTCYRAN,GTCYRDP,GTCMNAN,GTCMNDP,GTCTDIA)

  GTWhtm = ""

  GTWAccion = "ORCALENDAR.ASP?FCIN=" & GTCIN & "&FCYRAN=" & GTCYRAN & "&FCYRDP=" & GTCYRDP & "&FCMNAN=" & GTCMNAN & "&FCMNDP=" & GTCMNDP & "&FCTDIA=" & GTCTDIA & " "

  GTWhtm = GTWhtm & " onclick="" " & chr(13)
  GTWhtm = GTWhtm & " var dxformw1 = document.getElementById('FRMCALENDAR'); " & chr(13)
  GTWhtm = GTWhtm & " dxformw1.action = '" & GTWAccion & "'; " & chr(13)
  GTWhtm = GTWhtm & " DXCDATO  = '" & GTCDAT & "';" & chr(13)
  GTWhtm = GTWhtm & " DXCFORM  = '" & GTCFRM & "';" & chr(13)

  if GTCSMT = "N" then
     GTWhtm = GTWhtm & " DXCSMT  = 'N';" & chr(13)
  end if

  if GTCFCH = "1"  or GTCFCH = "2" then
     GTWhtm = GTWhtm & " DXCFCH     = '" & GTCFCH & "';" & chr(13)
     GTWhtm = GTWhtm & " DXCFCHDAT  = '" & GTCFCHDAT & "';" & chr(13)
  end if



  GTWhtm = GTWhtm & " dxformw1.submit(); " & chr(13)
  GTWhtm = GTWhtm & " document.getElementById('DXCALENDAR').style.display='';" & chr(13)
  GTWhtm = GTWhtm & " return false;" & chr(13)
  GTWhtm = GTWhtm & """ " & chr(13)

  GTCSETCALR = GTWhtm

end function

function GTCCALENDAR(GTCTRG)

  if GTCTRG = "" then
     GTCTRG = "fesquemas"
  elseif GTCTRG = "M" then
     GTCTRG = "fesquemas"
  elseif GTCTRG = "D" then
     GTCTRG = "WK4"
  end if

  GTWhtm = ""


  GTWhtm = GTWhtm & "<div id='DXCALENDAR' style='margin-top:1;margin-left:1;position:absolute;left:400px;top:200px;display:none;"
  GTWhtm = GTWhtm & "width:400;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"
 '''''' GTWhtm = GTWhtm & TABLA
  GTWhtm = GTWhtm & "</div>"

  GTWAccion = "XXXX"
  GTWhtm = GTWhtm & ofv.genform("FRMCALENDAR","",GTWAccion,GTCTRG,"post","","","si")



  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)

  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & " var DXCDATO  = '';" & chr(13)
  GTWhtm = GTWhtm & " var DXCFORM  = '';" & chr(13)
  GTWhtm = GTWhtm & " var DXCSMT   = 'S';" & chr(13)
  GTWhtm = GTWhtm & " var DXCFCH   = 'N';" & chr(13)
  GTWhtm = GTWhtm & " var DXCFCHDAT  = '';" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function DXCALDATE(DXCDATE,DXCDATEL,DXCDATES) {" & chr(13)
  GTWhtm = GTWhtm & "    "
  GTWhtm = GTWhtm & "   var dxformw   = document.getElementById(DXCFORM);" & chr(13)
  GTWhtm = GTWhtm & "   var dxdatow   =  document.getElementById(DXCDATO);" & chr(13)
''' ' GTWhtm = GTWhtm & "   var dxdatow   = document.c(DXCDATO);" & chr(13)
 GTWhtm = GTWhtm & " dxdatow.value=DXCDATE;" & chr(13) 'alert(dxdatow.value);'

  GTWhtm = GTWhtm & "  if (DXCFCH=='N') {" & chr(13)
  GTWhtm = GTWhtm & "     } else {" & chr(13)
  GTWhtm = GTWhtm & "      var dxcellw   = document.getElementById(DXCFCHDAT);" & chr(13)
  GTWhtm = GTWhtm & "      if (DXCFCH=='1') { " & chr(13)
  GTWhtm = GTWhtm & "          dxcellw.innerHTML=DXCDATEL;" & chr(13)
  GTWhtm = GTWhtm & "    } else {" & chr(13)
  GTWhtm = GTWhtm & "          dxcellw.innerHTML=DXCDATES;" & chr(13)
  GTWhtm = GTWhtm & "               }" & chr(13)
  GTWhtm = GTWhtm & "                   }" & chr(13)

  GTWhtm = GTWhtm & "  if (DXCSMT=='S') {dxformw.submit();}" & chr(13)
  GTWhtm = GTWhtm & "  else {document.getElementById('DXCALENDAR').style.display='none';} " & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)

  GTWhtm = GTWhtm & "</SCRIPT>"

  GTCCALENDAR = GTWhtm


end function

function GTCOBS()

  GTWhtm = ""

  GTWhtm = GTWhtm & "<div id='DX4' style='position:absolute;left:150px;top:250px;display:none;"
  GTWhtm = GTWhtm & "width:500px;height:350px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"

  GTWhtm = GTWhtm & "&nbsp;"

  GTWhtm = GTWhtm & "</div>"

  'GTWhtm = GTWhtm & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>"
   GTWhtm= GTWhtm &  "<script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & chr(13)
  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSWM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='800px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSSS() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='500px';" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='350px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSHM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='600px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)

  GTWhtm = GTWhtm & "function GTCOBSCLEX(e) {" & chr(13)
  GTWhtm = GTWhtm & " if(e != 'S')  { alert('Por favor, complete los campos.');}  else { document.getElementById('DX4').style.display='none'; } " & chr(13)

  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSCL() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.display='none';" & chr(13)

  GTWhtm = GTWhtm & " } " & chr(13)
  if session("yider")<> "" then
  session("yider")=replace(session("yider"),"*" ,"")
  ayider=split(session("yider")," ")
  if isarray(ayider) then
  zdelim= ""
  pals=""
   for ia=0 to ubound(ayider)
      pal= trim(ayider(ia) & "")
      if len(pal) > 3 then
      pals=pals & zdelim & pal
      zdelim= " "
      end if
   next
   session("yider")=pals
  end if
  GTWhtm = GTWhtm & " localSearchHighlight('?h="& session("yider")& "'); var z= document.getElementById('sxxxs');" & chr(13)
  GTWhtm = GTWhtm & "z.scrollIntoView(true);" & chr(13)
 end if
  GTWhtm = GTWhtm & "</SCRIPT>" '& pals & "222222222222222222"

  GTCOBS = GTWhtm


end function

function GTCOBSRSK()

  GTWhtm = ""

  GTWhtm = GTWhtm & "<div id='DXA' style='position:absolute;left:50px;top:250px;display:none;"
  GTWhtm = GTWhtm & "width:500px;height:350px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"

  GTWhtm = GTWhtm & "&nbsp;"

  GTWhtm = GTWhtm & "</div>"

  GTWhtm = GTWhtm & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>"

  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSWM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.width='800px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSSS() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.width='500px';" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.height='350px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSHM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.height='600px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCOBSCL() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DXA').style.display='none';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "</SCRIPT>"

  GTCOBSRSK = GTWhtm


end function

function GTCPTOFUN()

  GTWhtm = ""

  GTWhtm = GTWhtm & "<div id='DX4' style='position:absolute;left:50px;top:250px;display:none;"
  GTWhtm = GTWhtm & "width:500px;height:400px;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
  GTWhtm = GTWhtm & "background-color:" & session("APLBGD") & ";"
  GTWhtm = GTWhtm & "' >"

  GTWhtm = GTWhtm & "&nbsp;"

  GTWhtm = GTWhtm & "</div>"

  GTWhtm = GTWhtm & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>"

  GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & chr(13)
  GTWhtm = GTWhtm & chr(13)
  GTWhtm = GTWhtm & "function GTCPTOFUNWM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='800px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCPTOFUNSS() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.width='500px';" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='400px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "function GTCPTOFUNHM() {" & chr(13)
  GTWhtm = GTWhtm & "   document.getElementById('DX4').style.height='600px';" & chr(13)
  GTWhtm = GTWhtm & " } " & chr(13)
  GTWhtm = GTWhtm & "</SCRIPT>"

  GTCPTOFUN = GTWhtm


end function

FUNCTION ValidarInput()
    dim sSval


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

  END FUNCTION


  Function strPad( strToPad, toLength, padWith, padWhere )
  '====
      'padWhere: 1=right, 2=left, 3=both
      strToPad = strToPad & ""
      dim padBoth
      padBoth = false

      if not isNumeric( toLength ) then

       else
           toLength =  toLength
      end if

      if padWith = "" then padWith = " "

      if not isNumeric( padWhere ) then
          padWhere = 1
       else
           padWhere =  padWhere
          if padWhere > 3 or padWhere < 1 then padWhere = 1
           if padWhere = 3 then
              padBoth = true
              padWhere = 1
          end if
      end if

      do until len( strtoPad ) >= toLength
          select case padWhere
           case 1
               strtoPad = strtoPad & padWith
           case 2
               strtoPad = padWith & strtoPad
          end select
          if padBoth then
              if padWhere = 1 then
                  padWhere = 2
               else
                   padWhere = 1
              end if
          end if
      loop

      if len( strToPad ) > toLength then
          if ( padBoth and padWhere = 1 ) or ( padWhere = 2 and not padBoth ) then '=== no carriage return here
              strToPad = right( strToPad, toLength )
          else
              strToPad = left( strToPad, toLength )
          end if
      end if

      strPad = strToPad

End Function


' devuelve la fecha inicio y fin de un trimestre

Function GetFecTrim(Trimestre,ano,fecini,fecfin)


Select Case Trimestre
Case 1
  ConvDate = Dateserial( ano,1,1)
Case 2
  ConvDate = Dateserial( ano,4,1)
Case 3
  ConvDate = Dateserial( ano,7,1)
Case 4
  ConvDate = Dateserial( ano,10,1)
End Select

fecini=ConvDate
fecfin = DateAdd("q", 1, fecini)
fecfin = DateAdd("d", -1, fecfin)


End Function




' devuelve el trimestre de una fecha

function GetTrimestre(fecha)

 dia=day(fecha)
 mes=month(fecha)
 ano=year(fecha)
 fecha=dateserial(ano,mes,dia)
 GetTrimestre=datepart("q",fecha)

end function

'------------------------------------------------------------

Function reemplaza(texto, patrn, replStr)
        Dim regEx
        Set regEx = New RegExp            ' Crea una expresión regular.
        regEx.Pattern = patrn            ' Establece el modelo.
        regEx.IgnoreCase = True            ' Hace distinción entre mayúsculas y minúsculas.
        reemplaza = regEx.Replace(texto, replStr)   ' Realiza el reemplazo.
        set regEx = nothing
End Function


'------------------------------------------------------------

function EliminarAcentos(TEXTO)

if len(texto) <> 0 then
      texto = reemplaza(texto,"[ÀÁÄáàä]","A")
      texto = reemplaza(texto,"[ÈÉËèéë]","E")
      texto = reemplaza(texto,"[ÌÍÏìíï]","I")
      texto = reemplaza(texto,"[ÒÓÖòóö]","O")
      texto = reemplaza(texto,"[ÙÚÜùúü]","U")
end if
  EliminarAcentos = texto
End Function


FUNCTION GTCSETCALREX(GTCFRM, GTCDAT, GTCSMT, GTCFCH, GTCFCHDAT, GTCIN, GTCYRAN, GTCYRDP, GTCMNAN, GTCMNDP, GTCTDIA,GTCINIEX,GTCFINEX,GTCDIANTEX,GTCDIADPEX)

   GTWhtm = ""

   GTWAccion = "ORCALENDAR.ASP?FCIN=" & GTCIN & "&FCYRAN=" & GTCYRAN & "&FCYRDP=" & GTCYRDP & "&FCMNAN=" & GTCMNAN & "&FCMNDP=" & GTCMNDP & "&FCTDIA=" & GTCTDIA '& ""
   GTWAccion = GTWAccion & "&FDIASINI=" & GTCINIEX & "&FFIN=" & GTCFINEX & "&DIASANT=" & GTCDIANTEX & "&DIASDP=" & GTCDIADPEX
   GTWhtm = GTWhtm & " onclick="" " & CHR(13)
   GTWhtm = GTWhtm & " var dxformw1 = document.getElementById('FRMCALENDAR'); " & CHR(13)
   GTWhtm = GTWhtm & " dxformw1.action = '" & GTWAccion & "'; " & CHR(13)
   GTWhtm = GTWhtm & " DXCDATO  = '" & GTCDAT & "';" & CHR(13)
   GTWhtm = GTWhtm & " DXCFORM  = '" & GTCFRM & "';" & CHR(13)

   IF GTCSMT = "N" THEN
      GTWhtm = GTWhtm & " DXCSMT  = 'N';" & CHR(13)
   END IF

   IF GTCFCH = "1" OR GTCFCH = "2" THEN
      GTWhtm = GTWhtm & " DXCFCH     = '" & GTCFCH & "';" & CHR(13)
      GTWhtm = GTWhtm & " DXCFCHDAT  = '" & GTCFCHDAT & "';" & CHR(13)
   END IF

   GTWhtm = GTWhtm & " dxformw1.submit(); " & CHR(13)
   GTWhtm = GTWhtm & " document.getElementById('DXCALENDAR').style.display='';" & CHR(13)
   GTWhtm = GTWhtm & " return false;" & CHR(13)
   GTWhtm = GTWhtm & """ " & CHR(13)

   GTCSETCALREX = GTWhtm

END FUNCTION

'------------------------------------------------------------
Function BUSCARCAR(texto, patrn)
        Dim regEx
        Set regEx = New RegExp            ' Crea una expresión regular.
        regEx.Pattern = patrn            ' Establece el modelo.
        regEx.IgnoreCase = True            ' Hace distinción entre mayúsculas y minúsculas.
        SET XBUSCAR = regEx.Execute(texto)   ' Realiza el reemplazo.
        BUSCARCAR=XBUSCAR.Count > 0
        set XBUSCAR = nothing
        set regEx = nothing
End Function


FUNCTION BUSCAEXCAR(TEXTO,PATRON)
   TEXTO=TEXTO &""
   IF PATRON="" THEN PATRON="["& chr(9)  & chr(34) & "'ÀÁÄáàäÈÉËèéëÌÍÏìíïÒÓÖòóöÙÚÜùúü]"

    BUSCAEXCAR=BUSCARCAR(TEXTO,PATRON)

END FUNCTION

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



    function armaperfiles(hash,cn,id)
     delim=""
     dim sqlq
	 HashPerfiles=" var  " & hash & "  =new Hashtable();" & chr(13)

	    sqlq="select a.descri,b.*,c.objid from Tipoperfil a "
		sqlq=sqlq & " inner JOIN ecoperfil b on( b.codperfil=a.tipoPERfil)"
        sqlq=sqlq & " inner join   objsector c on(c.codsector=b.codsector and c.codobjeto="&id&")"
       ' response.write sqlq
	      set xpf=ofv.crearconsultaEx(sqlq,cn,1,parametros)
	     ' response.write sql
	      do while not xpf.eof
	      HashPerfiles= HashPerfiles & hash & ".put('pf" & xpf(1)	&"','" & xpf(0) &"'); " &  chr(13)
	      sperfiles= sperfiles & delim & "pf" & xpf(1) & "#" & xpf(0) & "#" & xpf(5)
	      delim="###"

	      xpf.movenext
	      loop
   call ofv.cerrarconsulta(xpf)
	armaperfiles=HashPerfiles
end function


function armasectores(hash,cn,id)
     delim=""
     dim sqlq
	 HashPerfiles=" var  " & hash & "  =new Hashtable();" & chr(13)

        delim=""
        sqlq=" Select b.descri,A.objid From objsector a "
		sqlq = sqlq & " inner join eco b on (b.eco=a.codsector)"

       sqlq =sqlq & " where a.codobjeto=0" & id
'response.write sqlq
	      set xpf=ofv.crearconsultaEx(sqlq,cn,1,parametros)
	     ' response.write sql
	      do while not xpf.eof
	      HashPerfiles= HashPerfiles & hash & ".put('idsec" & xpf(1)	&"','" & xpf(0) &"'); " &  chr(13)
	     ssectores= ssectores & delim & "idsec" & xpf(1) & "#" & xpf(0)
	      delim="###"

	      xpf.movenext
	      loop
   call ofv.cerrarconsulta(xpf)
	armasectores=HashPerfiles
end function


function retnum(s)

  for i =1 to len(s)

    if isnumeric(mid(s,i,1)) then
    retnum=retnum & (mid(s,i,1))
    end if

  next

end function

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

        StrSql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
        Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
        If Not RSM.EOF Then

            tipo=RSM(2)


            IF session("APLMDCLI") ="BC" THEN
            sTipobjeto1 = (RSM(1)) & ":&nbsp;"
            sTipobjeto2 = RSM(3) & ""
            sTipobjeto3 = " - " & (RSM(2) & "")
      sTipobjeto4 = " (" & (RSM(5) & "") & " "
            sTipobjeto5 = (RSM(6) & "") & ") "
            ELSE
            sTipobjeto1 = (RSM(1)) & ":&nbsp;"
            sTipobjeto2 = (RSM(3) & "")

            sTipobjeto3 = " - " & (RSM(2) & "")
            sTipobjeto4 = " (" & (RSM(5) & "")
            if ucase(RSM(5) & "")="O" THEN
             sTipobjeto5 = ")"
            ELSE
            sTipobjeto5 =  " " & (RSM(6) & "") & ") "
            END IF
            END IF
            If session("FULLDOCN") = "1" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2
            ElseIf session("FULLDOCN") = "2" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
            Elseif session("FULLDOCN") = "4" then
                sTipobjetoX =  sTipobjeto1 & tipo & " - " & sTipobjeto2 &  sTipobjeto4 & sTipobjeto5
            Else
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3  & sTipobjeto4 & sTipobjeto5
            End If



        End If

    End If





    Call ofv.cerrarconn(xcn)


    EsqnomPx = sTipobjetoX


End Function
%>