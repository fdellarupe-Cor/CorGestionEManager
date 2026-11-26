<!-- #INCLUDE FILE="GetinfoAD.inc" -->


<%

Dim arrr (2000,1)



' esta url es donde se encuentra el fileserver con los asp para leer directorios ,archivos, y leer files
'session("UrlGetdata")="http://server-20:80/coresv3/FileServer/asp/Getdata.asp"
' esto es para el upload

If SESSION("URLPORT")  = "Y" Then
	session("UrlGetdata")="http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") &"/coremanager/FileServer/asp/Getdata.asp"
Else
	session("UrlGetdata")="http://" & request.servervariables("SERVER_NAME") &"/coremanager/FileServer/asp/Getdata.asp"
End If



' este el el sito web origen
'session ("SiteCor")= "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT")
session ("SiteCor")= "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT")

Function callUrl(ROOT,url)
	Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")


	'url=session("UrlGetdata") &"?OP=RFILES&folder=" & root
	xmlHttp.Open "GET", url, False
	xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
	xmlHttp.Send()
	ss=xmlHttp.ResponseText
	callUrl = ss


End Function


Function getdata( XARCHIVO,op,folderx,NEWFOLDER)

	Response.Charset="ISO-8859-1"

	ruta=server.mappath(request.servervariables("script_name"))

	pos=InStrRev(ruta,"\")
	ruta=Mid(ruta,1,pos)

	If OP = "RF" And SESSION("FILETIPO") = "R" Then

		ruta = session ("DirUpload")

	End If
	ruta = session ("DirUpload")
	'' response.write "22222222222222" & ruta & "-" & newfolder &  "-" & folderx &  "-" & op  &  "-DIR:" &  ruta  & folderx  & "<BR>"

	Set fso = SERVER.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(ruta  & folderx) Then
		'response.write ruta  & folderx
		Set wwFolder = fso.getfolder(ruta  & folderx)

		'' IF op <> "RFILES" THEN op=""

		''  op=""

		Select Case OP

			Case "RF"
			'response.write ruta & folderx & OP
			Set wwsfolder = wwFolder.SubFolders

			For Each Fcc In wwsfolder
				getdata=getdata &  fcc.name & "|"
			Next


			Case "AF"

			Set wwsfolder = wwFolder.SubFolders
			wwsFolder.add(NEWFOLDER)


			Case "RFILES"

			Set Files = wwFolder.Files

			For Each Fcc In files
				'   response.write  f.name &"###" & f.DateCreated  & "|"
				getdata=getdata &  fcc.name &"###" & fcc.DateCreated  & "|"
			Next

			Case "F"

			Xarchivo=RUTA & "CARPETASCMD\" & XARCHIVO
			' RESPONSE.WRITE XARCHIVO
			If fso.FileExists(Trim(Xarchivo)) Then
				Set fff1 = fso.openTextFile(Trim(Xarchivo), 1, -1)
				Do While fff1.AtEndOfstream <> True
					getdata=getdata & fff1.readline
				Loop
				fff1.close
			End If

			Case "F1"


			If fso.FileExists(Trim(Xarchivo)) Then
				Set fff1 = fso.openTextFile(Trim(Xarchivo), 1, -1)
				Do While fff1.AtEndOfstream <> True
					getdata=getdata & fff1.readline
				Loop
				fff1.close
			End If
			Case "FSP"

			Xarchivo = RUTA & folderx & "\" & XARCHIVO
			If fso.FILEEXISTS(Trim(Xarchivo)) Then
				Set fff1 = fso.OPENTEXTFILE(Trim(Xarchivo), 1, -1)
				Do While fff1.ATENDOFSTREAM <> True
					getdata = getdata & fff1.READLINE & Chr(9)
				Loop
				fff1.CLOSE
			End If

			Case "FMOV"

			Xarchivo1 = RUTA & folderx & "\" & XARCHIVO
			Xarchivo2 = RUTA & NEWFOLDER & "\" & XARCHIVO

			''' RESPONSE.WRITE Xarchivo1 & " - A - " & Xarchivo2 & "<BR>"

			If fso.FILEEXISTS(Trim(Xarchivo1)) Then
				Set aFile = fso.GETFILE(Trim(Xarchivo1))
				aFile.MOVE Trim(Xarchivo2)
			End If
		End Select

	End If
End Function


Function delfile(arch)


	If  SESSION("FILETIPO") = "R" Then
		Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		DataToSend = "ARCH=" & arch
		url=session("UrlGetdata") & "?OP=DF"
		xmlHttp.Open "POST", url, False
		xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
		xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

		xmlHttp.Send(DataToSend)
		'RESPONSE.WRITE  xmlHttp.ResponseText
	Else
		ruta=server.mappath(request.servervariables("script_name"))

		pos=InStrRev(ruta,"\")
		ruta=Mid(ruta,1,pos) & arch


		Set fso = SERVER.CreateObject("Scripting.FileSystemObject")
		Set fp = fso.GetFile(ruta)

		If ((fp.attributes And 2) = 0) And  ((fp.attributes And 1)=0) Then
			fso.DeleteFile(ruta)
		End If

		Set fso=Nothing
	End If


End Function

Function callUrl(ROOT,url)
	Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")


	url=session("UrlGetdata") &"?OP=RFILES&folder=" & root
	xmlHttp.Open "GET", url, False
	xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
	xmlHttp.send
	ss=xmlHttp.ResponseText
	callUrl = ss


End Function


Function getfolders(root,url)


	'response.write SESSION("FILETIPO")
	Select Case SESSION("FILETIPO")

		Case "R"
		Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		url=session("UrlGetdata") & "?OP=RF&folder=" & root
		' response.write url

		xmlHttp.Open "GET", url, False
		'  xmlHTTP.setRequestHeader "AUTH_TYPE", "Negotiate"
		xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
		XMLHttp.setRequestHeader "charset", "ISO-8859-1"

		xmlHttp.Send()
		ss=xmlHttp.responseBody 'xmlHttp.ResponseText

		getfolders =  BinaryToString1(ss, "UTF-8")
		' response.write getfolders

		Case "L"
		'function getdata( XARCHIVO,op,folder,NEWFOLDER)
		getfolders=getdata("","RF",ROOT,"")
	End Select


End Function

Function getfilesinfolder(ROOT,url)
	'response.write SESSION("FILETIPO")
	Select Case SESSION("FILETIPO")

		Case  "R"

		Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")


		url=session("UrlGetdata") &"?OP=RFILES&folder=" & root
		xmlHttp.Open "GET", url, False
		xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
		XMLHttp.setRequestHeader "charset", "ISO-8859-1"
		xmlHttp.Send()
		ss=XMLHttp.responseBody  'xmlHttp.ResponseText

		getfilesinfolder =   BinaryToString1(ss, "UTF-8")


		Case "L"
		getfilesinfolder = getdata("","RFILES",root,"")
	End Select


End Function

Function getReadfile(ROOT,url)


	Select Case SESSION("FILETIPO")

		Case  "R"

		Set xmlHttp4 = Server.CreateObject("Msxml2.ServerXMLHTTP")


		urlE=session("UrlGetdata") &"?OP=F&xarchivo=" & root


		xmlHttp4.Open "GET", urlE, False
		xmlHttp4.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"


		xmlHttp4.Send()
		ss=xmlHttp4.ResponseText
		getReadfile = ss '& session("sitecor")

		Case "L"

		getReadfile=getdata(root,"F","","")
	End Select


End Function


Function getReadfile2(ROOT,url)

	Select Case SESSION("FILETIPO")

		Case  "R"

		Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")


		url2=session("UrlGetdata") &"?OP=F1&xarchivo=" & root & "&newfolder="& url

		xmlHttp.Open "GET", url2, False
		xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
		xmlHttp.Send()
		ss=xmlHttp.ResponseText
		getReadfile2 = ss '& session("sitecor")

		Case "L"

		getReadfile2=getdata(root,"F1","","")
	End Select


End Function



Function ADDFOLDER(root,newfolder,url)

	Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	url=session("UrlGetdata") & "?OP=AF&folder=" & root & "&NEWFOLDER="&NEWFOLDER
	xmlHttp.Open "GET", url, False
	xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
	xmlHttp.Send()
	'getfilesinfolder = xmlHttp.ResponseText

End Function




Function UserLogon()
	User=REQUEST.servervariables("LOGON_USER")
	User=Mid(User,InStr(1,User,"\")+1)
	UserLogon=User
	'UserLogon="usu1"
End Function


Function UserAuth()
	User=REQUEST.servervariables("AUTH_USER")
	User=Mid(User,InStr(1,User,"\")+1)
	UserAuth=User
End Function


Function getInfoECO(XGU_ECO)
	data = ""

	GTUcase = session("ECOPRV")

	If GTUCASE = "" Then GTUCASE = "I"


	Select Case GTUcase

		Case "I"

		If IsNumeric(XGU_ECO) Then
			WXGU_ECO = XGU_ECO
		Else
			WXGU_ECO = 1
		End If

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ECO,DESCRI,CODINTER  From ECO WHERE ECO = 0" & WXGU_ECO
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then
			session("SGU_SECO")      = xrs(0)
			session("SGU_SDESCRI")   = xrs(1)
			session("SGU_SCODINT")	 = xrs(2) & ""
		Else
			session("SGU_SECO")      = XGU_ECO
			session("SGU_SDESCRI")   = "NO"
			session("SGU_SCODINT")	 = "NO"

		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"

		session("SGU_SECO")      = XGU_ECO
		session("SGU_SDESCRI")   = "NO"
		session("SGU_SCODINT")	 = "NO"



	End Select

End Function

'-----------------------------------------
Function getInfouserPst(Legajo,IDPST,tipo)
	data = ""

	GTUcase = session("USRPRV")


	Select Case GTUcase

		Case "I"

		Call   getInfouserPstBcEX(Legajo,IDPST,tipo)
		Case "BC"

		Call   getInfouserPstBcEX(Legajo,IDPST,tipo)


		Case "BANELCO"

		Call   getInfouserPstBANELCOEX(Legajo,IDPST,tipo)

		Case "FINANSUR"

		Call   getInfouserPstBANELCOEX(Legajo,IDPST,tipo)

	End Select
End Function
'--------------------------------------------------------------------------------------------------



Function getInfoUser(Legajo)
	data = ""

	GTUcase = session("USRPRV")


	Select Case GTUcase

		Case "I"

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ID,descripcion,email,sector, puesto,site  From USUARIOS WHERE ID = '" & Legajo & "' "
		' response.write xStrSql
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then
			session("SGU_userid")      = xrs(0)
			session("SGU_nombre")      = xrs(1)
			session("SGU_mail")	   = xrs(2) & ""
			session("SGU_descrisector")= "MySect"
			session("SGU_idcostcenter")= xrs(3)
			session("SGU_descripuesto")= "MyPst"
			session("SGU_codigopuesto")= xrs(4)
			session("sgu_site")= xrs(5)
		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"
			session("sgu_site")= 1
		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"


		' DATOS DE LOS WEBSERVICES
		SoapAction = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getUserWithoutManagers"
		Method = "getUserWithoutManagers" ' METODO DEL WEBSERVICE A UTILIZAR

		END_POINT_URL = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
		CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"


		'Legajo = request.querystring("Legajo")
		If Legajo<>"" Then
			Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
			Connector.Property("EndPointURL") = END_POINT_URL
			Connector.Connect

			Connector.Property("SoapAction") = SoapAction & Method
			Connector.BeginMessage

			Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
			Serializer.Init Connector.InputStream

			Serializer.StartEnvelope
			Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
			Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
			Serializer.StartBody
			Serializer.StartElement Method, CALC_NS
			Serializer.StartElement "userId"
			Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
			Serializer.EndElement
			Serializer.EndElement
			Serializer.EndBody
			Serializer.EndEnvelope

			Connector.EndMessage

			Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
			Reader.Load Connector.OutputStream
			' OBTENCION DE LOS DATOS
			On Error Resume Next

		'	Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
		'	Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
		'	Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text
		set xtnode=Reader.Body.selectSingleNode("//multiRef/userId")
					if xtnode is nothing then
					 Userid=Legajo
					else
					 Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
					end if

					'Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text

					set xtnode=Reader.Body.selectSingleNode("//multiRef/displayName")

					if xtnode is nothing then
					 Nombre=Userid
					else
					 Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
					end if

					set xtnode=Reader.Body.selectSingleNode("//multiRef/workEMail")

					if xtnode is nothing then
					  Mail="NO@mail.com"
					else
					 Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text
            end if

			Set xml=Reader.Body
			Set raiz=xml.selectSingleNode("//getUserWithoutManagersResponse/getUserWithoutManagersReturn")

			idUserbg="id0" 'replace(raiz.getAttribute("href"),"#","")


			' tomo el idref de la persona
			Set nList = xml.selectNodes("//multiRef[@id='"& idUserbg &"']/person")

			For Each n In nList
				idpersona = Replace(n.getAttribute("href"),"#","")
			Next

			' tomo el idref del costcenter
			Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/costCenter")

			For Each n In nList
				idcostCenter = Replace(n.getAttribute("href"),"#","")
			Next


			' tomo el descripcion del costcenter /sector
			Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterDescription")
			For Each n In nList
				descrisector= n.text ' replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
			Next

			' tomo el id del costcenter
			Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterCode")
			For Each n In nList
				idcostcenter= n.text 'replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
			Next

			' tomo el id del puesto
			Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/functionTitle")
			For Each n In nList
				idpuesto= Replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
			Next

			'descripcion del puesto
			Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/titleDescription")
			For Each n In nList
				descripuesto=n.text   '  n.nodename &"=" & n.text & "<br>"
			Next


			'busco el codpuesto
			Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/comp_id")
			For Each n In nList
				idcomplemento= Replace(n.getAttribute("href"),"#","")
			Next

			' tomo el id del puesto

			Set nList = xml.selectNodes("//multiRef[@id='"&idcomplemento&"']/titleCode")
			For Each n In nList
				codigopuesto= n.text
			Next

			Set Reader=Nothing
			Set Serializer=Nothing
			Set Connector=Nothing
			'data= userid & "<br>"
			'data=data &  nombre & "<br>"
			'data=data &  mail & "<br>"
			'data=data &  descrisector & "<br>"
			'data=data &  idcostcenter & "<br>"
			'data=data &  descripuesto & "<br>"
			'data=data &  codigopuesto & "<br>"

			'getInfoUser=data
			session("SGU_userid")       = userid
			session("SGU_nombre")       = nombre
			session("SGU_mail")	    = mail
			session("SGU_descrisector") = descrisector
			session("SGU_idcostcenter") = idcostcenter
			session("SGU_descripuesto") = descripuesto
			session("SGU_codigopuesto") = codigopuesto

		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"


		End If



		Case "P" 'patagonia

		session("SGU_userid")      = Legajo
		session("SGU_nombre")      = ""
		session("SGU_mail")	   =  Legajo & "@bancopatagonia.com.ar"
		'session("SGU_descrisector")= "NO"
		'session("SGU_idcostcenter")= "0"
		'session("SGU_descripuesto")= "NO"
		'session("SGU_codigopuesto")= "NO"
		Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")
		Dirc=Dirc & "CorEssDll\"
		Dirc=Dirc & "NSBTest_2_50_05.xml" ' busqueda por texto o id
		'response.write  Dirc
		Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")
		xml4=""
		'response.write dirc
		xml4=leerArc(dirc)
		f=Now()
		' xml4=REPLACE(XML4,"XDRQCLUSR",Session("Usuario"))
		xml4=Replace(XML4,"XDRQCLUSR",Legajo )

		xml4=Replace(XML4,"XFECHACOR",fechaX(F))
		xml4=Replace(XML4,"XHORACOR", horaMinSec(F))

		xml4=Replace(XML4,"XXXDATO1",Legajo)  ' SE usa para buscar por id
		xml4=Replace(XML4,"XXXDATO","") ' SE usa para buscar por texto
		'    response.write xml4 & "<br>" & Legajo
		xmlHttp2.open "POST", session("GetInfo"), False
		xmlHttp2.send (xml4)
		Set xmldoc=xmlHTTP2.responsexml

		' response.write xmlHTTP2.responsetext

		Set z=xmldoc.selectSingleNode("//Evalls")

		Set nList = z.selectNodes("item")

		For Each n In nList

			If  N.hasChildNodes  Then 'primera ocurrencia tiene la info
				' response.write n.text
				xmldoc.LOADXML(N.text)

				Set nListx = XMLDOC.selectNodes("DATOS/USUARIO")
				For Each nx In nListx
					Set z1=nX.selectSingleNode("Codigo")
					Set z2=  nX.selectSingleNode("Nombre")
					Set z3=nX.selectSingleNode("Mail")
					session("SGU_userid")	= Z1.text
					session("SGU_nombre") 	= Z2.text
					session("SGU_mail")	= Legajo & "@bancopatagonia.com.ar"

				Next

				Exit For


			End If
		Next

		Set xmlHTTP2=Nothing
		Set xmldoc  =Nothing
		Set z	   =Nothing

		session("SGU_descrisector")= "NO"
		session("SGU_idcostcenter")= "0"
		session("SGU_descripuesto")= "NO"
		session("SGU_codigopuesto")= "NO"

		Case "T"

		getinfoUserAD(legajo)

		Case "S"

		getinfoUserAD(legajo)

		Case "BC"


		getinfoUserADBC(legajo)

		Case "GROI"

		getinfoUserADBCRED(legajo)
		Case  "BCOLU"
		getinfoUserAD(legajo)

		Case "LACAJA"

		getinfoUserADLACAJA(legajo)

		Case "CFA"

		getinfoUserAD(legajo)

		Case "RIVA"

		InfoDelLegajo(legajo)

		Case "BANELCO"
		getinfoUserADBanelco(LEGAJO)

		Case "FINANSUR"
		getinfoUserAD_FINANSUR(LEGAJO)
		Case "BPN"
		getinfoUserAD_BPN(LEGAJO)

		Case "GP"
		getinfoUserAD_GP(LEGAJO)
	End Select

End Function

Function getInfoUserEX(Legajo,DATOSE)
	data = ""

	GTUcase = session("USRPRV")


	Select Case GTUcase

		Case "I"

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ID,descripcion,email,sector, puesto,site  From USUARIOS WHERE ID = '" & Legajo & "' "
		' response.write xStrSql
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then
			session("SGU_userid")      = xrs(0)
			session("SGU_nombre")      = xrs(1)
			session("SGU_mail")	   = xrs(2) & ""
			session("SGU_descrisector")= "MySect"
			session("SGU_idcostcenter")= xrs(3)
			session("SGU_descripuesto")= "MyPst"
			session("SGU_codigopuesto")= xrs(4)
			session("sgu_site")= xrs(5)
		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"
			session("sgu_site")= 1
		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"


		' DATOS DE LOS WEBSERVICES
		SoapAction = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getUserWithoutManagers"
		Method = "getUserWithoutManagers" ' METODO DEL WEBSERVICE A UTILIZAR

		END_POINT_URL = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
		CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"


		'Legajo = request.querystring("Legajo")
		If Legajo<>"" Then
			Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
			Connector.Property("EndPointURL") = END_POINT_URL
			Connector.Connect

			Connector.Property("SoapAction") = SoapAction & Method
			Connector.BeginMessage

			Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
			Serializer.Init Connector.InputStream

			Serializer.StartEnvelope
			Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
			Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
			Serializer.StartBody
			Serializer.StartElement Method, CALC_NS
			Serializer.StartElement "userId"
			Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
			Serializer.EndElement
			Serializer.EndElement
			Serializer.EndBody
			Serializer.EndEnvelope

			Connector.EndMessage

			Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
			Reader.Load Connector.OutputStream
			' OBTENCION DE LOS DATOS


		'	Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
		'	Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
		'	Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text
		set xtnode=Reader.Body.selectSingleNode("//multiRef/userId")
					if xtnode is nothing then
					 Userid=Legajo
					else
					 Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
					end if

					'Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text

					set xtnode=Reader.Body.selectSingleNode("//multiRef/displayName")

					if xtnode is nothing then
					 Nombre=Userid
					else
					 Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
					end if

					set xtnode=Reader.Body.selectSingleNode("//multiRef/workEMail")

					if xtnode is nothing then
					  Mail="NO@mail.com"
					else
					 Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text
                    end if

			Set xml=Reader.Body
			'Set raiz=xml.selectSingleNode("//getUserWithoutManagersResponse/getUserWithoutManagersReturn")

			'id=Replace(raiz.getAttribute("href"),"#","")
            idUserbg="id0" 'replace(raiz.getAttribute("href"),"#","")

			' tomo el idref de la persona
			Set nList = xml.selectNodes("//multiRef[@id='"& idUserbg &"']/person")

			For Each n In nList
				idpersona = Replace(n.getAttribute("href"),"#","")
			Next

			' tomo el idref del costcenter
			Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/costCenter")

			For Each n In nList
				idcostCenter = Replace(n.getAttribute("href"),"#","")
			Next


			' tomo el descripcion del costcenter /sector
			Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterDescription")
			For Each n In nList
				descrisector= n.text ' replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
			Next

			' tomo el id del costcenter
			Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterCode")
			For Each n In nList
				idcostcenter= n.text 'replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
			Next

			' tomo el id del puesto
			Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/functionTitle")
			For Each n In nList
				idpuesto= Replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
			Next

			'descripcion del puesto
			Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/titleDescription")
			For Each n In nList
				descripuesto=n.text   '  n.nodename &"=" & n.text & "<br>"
			Next


			'busco el codpuesto
			Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/comp_id")
			For Each n In nList
				idcomplemento= Replace(n.getAttribute("href"),"#","")
			Next

			' tomo el id del puesto

			Set nList = xml.selectNodes("//multiRef[@id='"&idcomplemento&"']/titleCode")
			For Each n In nList
				codigopuesto= n.text
			Next

			Set Reader=Nothing
			Set Serializer=Nothing
			Set Connector=Nothing
			'data= userid & "<br>"
			'data=data &  nombre & "<br>"
			'data=data &  mail & "<br>"
			'data=data &  descrisector & "<br>"
			'data=data &  idcostcenter & "<br>"
			'data=data &  descripuesto & "<br>"
			'data=data &  codigopuesto & "<br>"

			'getInfoUser=data
			session("SGU_userid")       = userid
			session("SGU_nombre")       = nombre
			session("SGU_mail")	    = mail
			session("SGU_descrisector") = descrisector
			session("SGU_idcostcenter") = idcostcenter
			session("SGU_descripuesto") = descripuesto
			session("SGU_codigopuesto") = codigopuesto

		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"


		End If



		Case "P" 'patagonia

		session("SGU_userid")      = Legajo
		session("SGU_nombre")      = ""
		session("SGU_mail")	   =  Legajo & "@bancopatagonia.com.ar"
		'session("SGU_descrisector")= "NO"
		'session("SGU_idcostcenter")= "0"
		'session("SGU_descripuesto")= "NO"
		'session("SGU_codigopuesto")= "NO"
		Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")
		Dirc=Dirc & "CorEssDll\"
		Dirc=Dirc & "NSBTest_2_50_05.xml" ' busqueda por texto o id
		'response.write  Dirc
		Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")
		xml4=""
		'response.write dirc
		xml4=leerArc(dirc)
		f=Now()
		' xml4=REPLACE(XML4,"XDRQCLUSR",Session("Usuario"))
		xml4=Replace(XML4,"XDRQCLUSR",Legajo )

		xml4=Replace(XML4,"XFECHACOR",fechaX(F))
		xml4=Replace(XML4,"XHORACOR", horaMinSec(F))

		xml4=Replace(XML4,"XXXDATO1",Legajo)  ' SE usa para buscar por id
		xml4=Replace(XML4,"XXXDATO","") ' SE usa para buscar por texto
		'    response.write xml4 & "<br>" & Legajo
		xmlHttp2.open "POST", session("GetInfo"), False
		xmlHttp2.send (xml4)
		Set xmldoc=xmlHTTP2.responsexml

		' response.write xmlHTTP2.responsetext

		Set z=xmldoc.selectSingleNode("//Evalls")

		Set nList = z.selectNodes("item")

		For Each n In nList

			If  N.hasChildNodes  Then 'primera ocurrencia tiene la info
				' response.write n.text
				xmldoc.LOADXML(N.text)

				Set nListx = XMLDOC.selectNodes("DATOS/USUARIO")
				For Each nx In nListx
					Set z1=nX.selectSingleNode("Codigo")
					Set z2=  nX.selectSingleNode("Nombre")
					Set z3=nX.selectSingleNode("Mail")
					session("SGU_userid")	= Z1.text
					session("SGU_nombre") 	= Z2.text
					session("SGU_mail")	= Legajo & "@bancopatagonia.com.ar"

				Next

				Exit For


			End If
		Next

		Set xmlHTTP2=Nothing
		Set xmldoc  =Nothing
		Set z	   =Nothing

		session("SGU_descrisector")= "NO"
		session("SGU_idcostcenter")= "0"
		session("SGU_descripuesto")= "NO"
		session("SGU_codigopuesto")= "NO"

		Case "T"

		getinfoUserAD(legajo)

		Case "S"

		getinfoUserAD(legajo)

		Case "BC"


		getinfoUserADBC(legajo)

		Case "GROI"

		getinfoUserADBCRED(legajo)
		Case  "BCOLU"
		getinfoUserAD(legajo)

		Case "LACAJA"

		getinfoUserADLACAJA(legajo)

		Case "CFA"

		getinfoUserAD(legajo)

		Case "RIVA"

		InfoDelLegajo(legajo)

		Case "BANELCO"
		getinfoUserADBanelco(LEGAJO)

		Case "FINANSUR"
		getinfoUserAD_FINANSUR(LEGAJO)
		Case "BPN"
		getinfoUserAD_BPN(LEGAJO)

		Case "GP"
		Call  getinfoUserADEX_GP(LEGAJO,DATOSE)
	End Select

End Function


Function getInfoUserw(Legajo)
	data = ""



	On Error Resume Next

	GTUcase = session("USRPRV")


	Select Case GTUcase

		Case "I"
		'  case "E"

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ID,descripcion,email,sector, puesto  From USUARIOS WHERE ID = '" & Legajo & "' "
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then
			session("SGU_userid")      = xrs(0)
			session("SGU_nombre")      = xrs(1)
			session("SGU_mail")	   = xrs(2) & ""
			session("SGU_descrisector")= "MySect"
			session("SGU_idcostcenter")= xrs(3)
			session("SGU_descripuesto")= "MyPst"
			session("SGU_codigopuesto")= xrs(4)
		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"
		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"
		' DATOS DE LOS WEBSERVICES
		SoapAction = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getUserWithoutManagers"
		Method = "getUserWithoutManagers" ' METODO DEL WEBSERVICE A UTILIZAR

		END_POINT_URL = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
		CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"


		'Legajo = request.querystring("Legajo")
		If Legajo<>"" Then
			Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
			Connector.Property("EndPointURL") = END_POINT_URL
			Connector.Property("Timeout")=9000
			Connector.Connect

			If Err.number = 0 Then

				Connector.Property("SoapAction") = SoapAction & Method
				Connector.BeginMessage

				Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
				Serializer.Init Connector.InputStream

				Serializer.StartEnvelope
				Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
				Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
				Serializer.StartBody
				Serializer.StartElement Method, CALC_NS
				Serializer.StartElement "userId"
				Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
				Serializer.EndElement
				Serializer.EndElement
				Serializer.EndBody
				Serializer.EndEnvelope

				Connector.EndMessage

			End If

			If Err.number = 0 Then


				Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
				Reader.Load Connector.OutputStream
				' OBTENCION DE LOS DATOS

			End If

			If Err.number = 0 Then


				'Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
				'Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
				'Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text
			 set xtnode=Reader.Body.selectSingleNode("//multiRef/userId")
			if xtnode is nothing then
			 Userid=Legajo
			else
			 Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
			end if

			'Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text

			set xtnode=Reader.Body.selectSingleNode("//multiRef/displayName")

			if xtnode is nothing then
			 Nombre=Userid
			else
			 Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
			end if

			set xtnode=Reader.Body.selectSingleNode("//multiRef/workEMail")

			if xtnode is nothing then
			  Mail="NO@mail.com"
			else
			 Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text
            end if

				Set xml=Reader.Body
				'set raiz=xml.selectSingleNode("//getUserWithoutManagersResponse/getUserWithoutManagersReturn")

				idUserbg="id0" 'replace(raiz.getAttribute("href"),"#","")


				' tomo el idref de la persona
				Set nList = xml.selectNodes("//multiRef[@id='"& idUserbg &"']/person")

				For Each n In nList
					idpersona = Replace(n.getAttribute("href"),"#","")
				Next

				' tomo el idref del costcenter
				Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/costCenter")

				For Each n In nList
					idcostCenter = Replace(n.getAttribute("href"),"#","")
				Next


				' tomo el descripcion del costcenter /sector
				Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterDescription")
				For Each n In nList
					descrisector= n.text ' replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
				Next

				' tomo el id del costcenter
				Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterCode")
				For Each n In nList
					idcostcenter= n.text 'replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
				Next

				' tomo el id del puesto
				Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/functionTitle")
				For Each n In nList
					idpuesto= Replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
				Next

				'descripcion del puesto
				Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/titleDescription")
				For Each n In nList
					descripuesto=n.text   '  n.nodename &"=" & n.text & "<br>"
				Next


				'busco el codpuesto
				Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/comp_id")
				For Each n In nList
					idcomplemento= Replace(n.getAttribute("href"),"#","")
				Next

				' tomo el id del puesto

				Set nList = xml.selectNodes("//multiRef[@id='"&idcomplemento&"']/titleCode")
				For Each n In nList
					codigopuesto= n.text
				Next

				Set Reader=Nothing
				Set Serializer=Nothing
				Set Connector=Nothing
				'data= userid & "<br>"
				'data=data &  nombre & "<br>"
				'data=data &  mail & "<br>"
				'data=data &  descrisector & "<br>"
				'data=data &  idcostcenter & "<br>"
				'data=data &  descripuesto & "<br>"
				'data=data &  codigopuesto & "<br>"

				'getInfoUser=data

			End If

			If Err.number <> 0 Then


				session("SGU_userid")      = Legajo
				session("SGU_nombre")      = "NO DISPONIBLE - " & Err.number & "&nbsp;" & Err.description
				session("SGU_mail")	   = "NO@disponible.com"
				session("SGU_descrisector")= "NO disponible"
				session("SGU_idcostcenter")= "0"
				session("SGU_descripuesto")= "NO disponible"
				session("SGU_codigopuesto")= "NO"


			Else


				session("SGU_userid")       = userid
				session("SGU_nombre")       = nombre
				session("SGU_mail")	    = mail
				session("SGU_descrisector") = descrisector
				session("SGU_idcostcenter") = idcostcenter
				session("SGU_descripuesto") = descripuesto
				session("SGU_codigopuesto") = codigopuesto


			End If
		End If


		Case "P" 'patagonia

		session("SGU_userid")      = Legajo
		session("SGU_nombre")      = ""
		session("SGU_mail")	   =  Legajo & "@bancopatagonia.com.ar"
		'session("SGU_descrisector")= "NO"
		'session("SGU_idcostcenter")= "0"
		'session("SGU_descripuesto")= "NO"
		'session("SGU_codigopuesto")= "NO"
		Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")
		Dirc=Dirc & "CorEssDll\"
		Dirc=Dirc & "NSBTest_2_50_05.xml" ' busqueda por texto o id
		'response.write  Dirc
		Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")
		xml4=""
		'response.write dirc
		xml4=leerArc(dirc)
		f=Now()
		' xml4=REPLACE(XML4,"XDRQCLUSR",Session("Usuario"))
		xml4=Replace(XML4,"XDRQCLUSR",Legajo )

		xml4=Replace(XML4,"XFECHACOR",fechaX(F))
		xml4=Replace(XML4,"XHORACOR", horaMinSec(F))

		xml4=Replace(XML4,"XXXDATO1",Legajo)  ' SE usa para buscar por id
		xml4=Replace(XML4,"XXXDATO","") ' SE usa para buscar por texto
		'    response.write xml4 & "<br>" & Legajo
		xmlHttp2.open "POST", session("GetInfo"), False
		xmlHttp2.send (xml4)
		Set xmldoc=xmlHTTP2.responsexml

		' response.write xmlHTTP2.responsetext

		Set z=xmldoc.selectSingleNode("//Evalls")

		Set nList = z.selectNodes("item")

		For Each n In nList

			If  N.hasChildNodes  Then 'primera ocurrencia tiene la info
				' response.write n.text
				xmldoc.LOADXML(N.text)

				Set nListx = XMLDOC.selectNodes("DATOS/USUARIO")
				For Each nx In nListx
					Set z1=nX.selectSingleNode("Codigo")
					Set z2=  nX.selectSingleNode("Nombre")
					Set z3=nX.selectSingleNode("Mail")
					session("SGU_userid")	= Z1.text
					session("SGU_nombre") 	= Z2.text
					session("SGU_mail")	= Legajo & "@bancopatagonia.com.ar"

				Next

				Exit For


			End If
		Next

		Set xmlHTTP2=Nothing
		Set xmldoc  =Nothing
		Set z	   =Nothing

		session("SGU_descrisector")= "NO"
		session("SGU_idcostcenter")= "0"
		session("SGU_descripuesto")= "NO"
		session("SGU_codigopuesto")= "NO"

		Case "CFA"

		getInfouserWAD(legajo)

		Case "T"

		getInfouserWAD(legajo)


		Case "S"
		getInfouserWSV(legajo)


		Case "BC"
		getInfouserWBC(legajo)

		Case "GROI"

		getinfoUserWBCRED(legajo)

		Case  "BCOLU"
		getInfouserWBCOLU(legajo)

		Case "LACAJA"

		getInfouserWLACAJA(legajo)

		Case "RIVA"

		InfoDelLegajoW(legajo)

		Case "BANELCO"

		getInfouserWBANELCO(legajo)

		Case "FINANSUR"

		getInfouserW_FINANSUR(legajo)

		Case "BPN"

		getInfouserW_BPN(legajo)

		Case "GP"

		getInfouserW_GP(legajo)

	End Select

End Function

Function getInfoUserwEX(Legajo,DATOSE)
	data = ""



	On Error Resume Next

	GTUcase = session("USRPRV")


	Select Case GTUcase

		Case "I"
		'  case "E"

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ID,descripcion,email,sector, puesto  From USUARIOS WHERE ID = '" & Legajo & "' "
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then
			session("SGU_userid")      = xrs(0)
			session("SGU_nombre")      = xrs(1)
			session("SGU_mail")	   = xrs(2) & ""
			session("SGU_descrisector")= "MySect"
			session("SGU_idcostcenter")= xrs(3)
			session("SGU_descripuesto")= "MyPst"
			session("SGU_codigopuesto")= xrs(4)
		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"
		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"
		' case "I"


		' DATOS DE LOS WEBSERVICES
		SoapAction = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getUserWithoutManagers"
		Method = "getUserWithoutManagers" ' METODO DEL WEBSERVICE A UTILIZAR

		END_POINT_URL = "http://ecompany-prod.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
		CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"


		'Legajo = request.querystring("Legajo")
		If Legajo<>"" Then
			Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
			Connector.Property("EndPointURL") = END_POINT_URL
			Connector.Connect

			If Err.number = 0 Then

				Connector.Property("SoapAction") = SoapAction & Method
				Connector.BeginMessage

				Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
				Serializer.Init Connector.InputStream

				Serializer.StartEnvelope
				Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
				Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
				Serializer.StartBody
				Serializer.StartElement Method, CALC_NS
				Serializer.StartElement "userId"
				Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
				Serializer.EndElement
				Serializer.EndElement
				Serializer.EndBody
				Serializer.EndEnvelope

				Connector.EndMessage

			End If

			If Err.number = 0 Then


				Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
				Reader.Load Connector.OutputStream
				' OBTENCION DE LOS DATOS

			End If

			If Err.number = 0 Then


			'	Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
			'	Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
			'	Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text

set xtnode=Reader.Body.selectSingleNode("//multiRef/userId")
			if xtnode is nothing then
			 Userid=Legajo
			else
			 Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text
			end if

			'Userid=Reader.Body.selectSingleNode("//multiRef/userId").Text

			set xtnode=Reader.Body.selectSingleNode("//multiRef/displayName")

			if xtnode is nothing then
			 Nombre=Userid
			else
			 Nombre=Reader.Body.selectSingleNode("//multiRef/displayName").Text
			end if

			set xtnode=Reader.Body.selectSingleNode("//multiRef/workEMail")

			if xtnode is nothing then
			  Mail="NO@mail.com"
			else
			 Mail=Reader.Body.selectSingleNode("//multiRef/workEMail").Text
            end if

				Set xml=Reader.Body
				Set raiz=xml.selectSingleNode("//getUserWithoutManagersResponse/getUserWithoutManagersReturn")

				iduserbg="id0" 'Replace(raiz.getAttribute("href"),"#","")


				' tomo el idref de la persona
				Set nList = xml.selectNodes("//multiRef[@id='"& iduserbg &"']/person")

				For Each n In nList
					idpersona = Replace(n.getAttribute("href"),"#","")
				Next

				' tomo el idref del costcenter
				Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/costCenter")

				For Each n In nList
					idcostCenter = Replace(n.getAttribute("href"),"#","")
				Next


				' tomo el descripcion del costcenter /sector
				Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterDescription")
				For Each n In nList
					descrisector= n.text ' replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
				Next

				' tomo el id del costcenter
				Set nList = xml.selectNodes("//multiRef[@id='"&idcostCenter&"']/costCenterCode")
				For Each n In nList
					idcostcenter= n.text 'replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
				Next

				' tomo el id del puesto
				Set nList = xml.selectNodes("//multiRef[@id='"&idpersona&"']/functionTitle")
				For Each n In nList
					idpuesto= Replace(n.getAttribute("href"),"#","") ' n.nodename &"=" & n.text & "<br>"
				Next

				'descripcion del puesto
				Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/titleDescription")
				For Each n In nList
					descripuesto=n.text   '  n.nodename &"=" & n.text & "<br>"
				Next


				'busco el codpuesto
				Set nList = xml.selectNodes("//multiRef[@id='"&idpuesto&"']/comp_id")
				For Each n In nList
					idcomplemento= Replace(n.getAttribute("href"),"#","")
				Next

				' tomo el id del puesto

				Set nList = xml.selectNodes("//multiRef[@id='"&idcomplemento&"']/titleCode")
				For Each n In nList
					codigopuesto= n.text
				Next

				Set Reader=Nothing
				Set Serializer=Nothing
				Set Connector=Nothing
				'data= userid & "<br>"
				'data=data &  nombre & "<br>"
				'data=data &  mail & "<br>"
				'data=data &  descrisector & "<br>"
				'data=data &  idcostcenter & "<br>"
				'data=data &  descripuesto & "<br>"
				'data=data &  codigopuesto & "<br>"

				'getInfoUser=data

			End If

			If Err.number <> 0 Then


				session("SGU_userid")      = Legajo
				session("SGU_nombre")      = "NO DISPONIBLE - " & Err.number & "&nbsp;" & Err.description
				session("SGU_mail")	   = "NO@disponible.com"
				session("SGU_descrisector")= "NO disponible"
				session("SGU_idcostcenter")= "0"
				session("SGU_descripuesto")= "NO disponible"
				session("SGU_codigopuesto")= "NO"


			Else


				session("SGU_userid")       = userid
				session("SGU_nombre")       = nombre
				session("SGU_mail")	    = mail
				session("SGU_descrisector") = descrisector
				session("SGU_idcostcenter") = idcostcenter
				session("SGU_descripuesto") = descripuesto
				session("SGU_codigopuesto") = codigopuesto


			End If







		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"


		End If


		Case "P" 'patagonia

		session("SGU_userid")      = Legajo
		session("SGU_nombre")      = ""
		session("SGU_mail")	   =  Legajo & "@bancopatagonia.com.ar"
		'session("SGU_descrisector")= "NO"
		'session("SGU_idcostcenter")= "0"
		'session("SGU_descripuesto")= "NO"
		'session("SGU_codigopuesto")= "NO"
		Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")
		Dirc=Dirc & "CorEssDll\"
		Dirc=Dirc & "NSBTest_2_50_05.xml" ' busqueda por texto o id
		'response.write  Dirc
		Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")
		xml4=""
		'response.write dirc
		xml4=leerArc(dirc)
		f=Now()
		' xml4=REPLACE(XML4,"XDRQCLUSR",Session("Usuario"))
		xml4=Replace(XML4,"XDRQCLUSR",Legajo )

		xml4=Replace(XML4,"XFECHACOR",fechaX(F))
		xml4=Replace(XML4,"XHORACOR", horaMinSec(F))

		xml4=Replace(XML4,"XXXDATO1",Legajo)  ' SE usa para buscar por id
		xml4=Replace(XML4,"XXXDATO","") ' SE usa para buscar por texto
		'    response.write xml4 & "<br>" & Legajo
		xmlHttp2.open "POST", session("GetInfo"), False
		xmlHttp2.send (xml4)
		Set xmldoc=xmlHTTP2.responsexml

		' response.write xmlHTTP2.responsetext

		Set z=xmldoc.selectSingleNode("//Evalls")

		Set nList = z.selectNodes("item")

		For Each n In nList

			If  N.hasChildNodes  Then 'primera ocurrencia tiene la info
				' response.write n.text
				xmldoc.LOADXML(N.text)

				Set nListx = XMLDOC.selectNodes("DATOS/USUARIO")
				For Each nx In nListx
					Set z1=nX.selectSingleNode("Codigo")
					Set z2=  nX.selectSingleNode("Nombre")
					Set z3=nX.selectSingleNode("Mail")
					session("SGU_userid")	= Z1.text
					session("SGU_nombre") 	= Z2.text
					session("SGU_mail")	= Legajo & "@bancopatagonia.com.ar"

				Next

				Exit For


			End If
		Next

		Set xmlHTTP2=Nothing
		Set xmldoc  =Nothing
		Set z	   =Nothing

		session("SGU_descrisector")= "NO"
		session("SGU_idcostcenter")= "0"
		session("SGU_descripuesto")= "NO"
		session("SGU_codigopuesto")= "NO"

		Case "CFA"

		getInfouserWAD(legajo)

		Case "T"

		getInfouserWAD(legajo)


		Case "S"
		getInfouserWSV(legajo)


		Case "BC"
		getInfouserWBC(legajo)

		Case "GROI"

		getinfoUserWBCRED(legajo)

		Case  "BCOLU"
		getInfouserWBCOLU(legajo)

		Case "LACAJA"

		getInfouserWLACAJA(legajo)

		Case "RIVA"

		InfoDelLegajoW(legajo)

		Case "BANELCO"

		getInfouserWBANELCO(legajo)

		Case "FINANSUR"

		getInfouserW_FINANSUR(legajo)

		Case "BPN"

		getInfouserW_BPN(legajo)

		Case "GP"

		Call  getInfouserWEX_GP(legajo,DATOSE)

	End Select

End Function
'--------------------------------------------------------------------------------------------------

Function getGroupsUser(Legajo)
	data = ""


	GTUcase = session("USRPRV")


	Select Case GTUcase

		Case "I"

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select grupo  From USUARIOS WHERE ID = '" & Legajo & "' "
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then
			session("SGU_grupos")       = xrs(0)

		Else
			session("SGU_grupos")       = "NO"

		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		'    session("SGU_grupos")


		Case "E"

		' DATOS DE LOS WEBSERVICES
		SoapAction = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getGroups"
		Method = "getGroups" ' METODO DEL WEBSERVICE A UTILIZAR

		END_POINT_URL = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
		CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"


		'Legajo = request.querystring("Legajo")
		If Legajo<>"" Then
			Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
			Connector.Property("EndPointURL") = END_POINT_URL
			Connector.Connect

			Connector.Property("SoapAction") = SoapAction & Method
			Connector.BeginMessage

			Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
			Serializer.Init Connector.InputStream

			Serializer.StartEnvelope
			Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
			Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
			Serializer.StartBody
			Serializer.StartElement Method, CALC_NS
			Serializer.StartElement "userId"
			Serializer.WriteString  Legajo  ' LEGAJO DEL USUARIO A BUSCAR
			Serializer.EndElement
			Serializer.EndElement
			Serializer.EndBody
			Serializer.EndEnvelope

			Connector.EndMessage

			Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
			Reader.Load Connector.OutputStream

			' prueba
			'response.write Reader.Body.xml
			Set nList = Reader.Body.selectNodes("//multiRef/name")
			session("SGU_grupos")="DEF" & Chr(9)
			Dim iLoop
			For Each n In nList
				If n.text <>"" Then
					'response.write  iloop & " grupo=" & n.text & "<br>"
					session("SGU_grupos")=session("SGU_grupos") & n.text & Chr(9)
				End If
				iloop = iloop + 1
			Next
			'---------------------------
			'data =  Reader.Body.selectSingleNode("//multiRef/userId").Text & chr(13)
			'data =  data & Reader.Body.selectSingleNode("//multiRef/displayName").Text & chr(13)
			'data =  data & Reader.Body.selectSingleNode("//multiRef/workEMail").Text & chr(13)
			'data =  data & Reader.Body.selectSingleNode("//multiRef/costCenterCode").Text & chr(13)
			'data =  data & Reader.Body.selectSingleNode("//multiRef/costCenterDescription").Text




			'getGroupsUser=data
			Set Reader=Nothing
			Set Serializer=Nothing
			Set Connector=Nothing


		End If

		Case "CFA"

		getGroupsUserADSV(Legajo)

		Case "T"

		getGroupsUserAD(Legajo)

		Case "S"

		getGroupsUserADSV(Legajo)

		Case "BC"

		getGroupsUserADSV(Legajo)

		Case "GROI"

		getGroupsUserBCRED(Legajo)
		Case "BCOLU"
		getGroupsUserADBCOLU(Legajo)

		Case "LACAJA"

		getGroupsUserADLaCaja(Legajo)



		Case "BANELCO"

		getGroupsUserADBANELCO(Legajo)

		Case "FINANSUR"

		getGroupsUserAD_FINANSUR(Legajo)
		Case "RIVA"
		GrupoDelLegajo(LEGAJO)

		Case "BPN"
		getGroupsUserAD_BPN(Legajo)

		Case "GP"
		getGroupsUserAD_GP(Legajo)

	End Select


End Function

'--------------------------------------------------------------------------------------------------

Function buscarGrupoAd(GRUPOAD)
	data = ""


	GTUcase = session("USRPRV")


	Select Case GTUcase


		Case "GP"
		buscarGrupoAd= buscarGrupoAdGP(GRUPOAD)

		Case Else
           if len(GRUPOAD & "") > 0 or GRUPOAD <> "" then
		XKEYGRP = Split(GRUPOAD,Chr(9))
		XDELIMG=""
		XName=""
		session("SGU_UIGuserid")=""
		session("SGU_UIGdescri")=""
		xgusep=""

		For XGUINX=0 To UBound(XKEYGRP)
			XName= XName & XDELIMG & "  '" & XKEYGRP(XGUINX) & "'"
			XDELIMG= " ,"
		Next

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ID,DESCRIPCION  From GRUPOS  WHERE ID IN ( " & XName & ") "
		'  RESPONSE.WRITE  xStrSql
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		Do While Not   xrs.EOF
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & xrs (1)
			xgusep = Chr(9)
			xrs.MOVENEXT
		Loop

        buscarGrupoAd=""
		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)

		end if

	End Select


End Function


Sub MuestraNodos ( Nodos )
	Dim oNodo
	For Each oNodo In Nodos
		response.write oNodo.nodeName & "=" & oNodo.nodeValue & "<br>"

		If oNodo.hasChildNodes Then
			response.write Chr(9)
			Call   MuestraNodos (oNodo.childNodes)
		End If


	Next

End Sub



Sub getUsrIngroup(grupo)
	Dim xml,root,i

	GTUcase = session("USRPRV")

	xgusep                   = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Select Case GTUcase

		Case "I"


		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ID,descripcion  From USUARIOS WHERE grupo = '" & grupo & "' order by descripcion "
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then

			Do While Not XRS.EOF

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
				session("SGU_UIGdescri") = session("SGU_UIGdescri") &  xgusep & replace(xrs(1) & "" , "'"," ")

				xgusep                   = Chr(9)

				XRS.MOVENEXT
			Loop

		Else

		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"

		' DATOS DE LOS WEBSERVICES
		SoapAction = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory/getMembersInGroup"
		Method = "getMembersInGroup" ' METODO DEL WEBSERVICE A UTILIZAR

		END_POINT_URL = "http://ecompany.bancogalicia.com.ar/DirectorioGlobal/services/GlobalDirectory?wsdl"
		CALC_NS = "http://schemas.xmlsoap.org/soap/envelope/"


		Set Connector = Server.CreateObject("MSSOAP.HttpConnector30")
		Connector.Property("EndPointURL") = END_POINT_URL
		Connector.Connect

		Connector.Property("SoapAction") = SoapAction & Method
		Connector.BeginMessage

		Set Serializer = Server.CreateObject("MSSOAP.SoapSerializer30")
		Serializer.Init Connector.InputStream

		Serializer.StartEnvelope
		Serializer.SoapAttribute "xmlns:xsi", , "http://www.w3.org/1999/XMLSchema-instance"
		Serializer.SoapAttribute "xmlns:xsd", , "http://www.w3.org/1999/XMLSchema"
		Serializer.StartBody
		Serializer.StartElement Method, CALC_NS
		Serializer.StartElement "cn"
		Serializer.WriteString  grupo  ' LEGAJO DEL USUARIO A BUSCAR
		Serializer.EndElement
		Serializer.EndElement
		Serializer.EndBody
		Serializer.EndEnvelope

		Connector.EndMessage

		Set Reader = Server.CreateObject("MSSOAP.SoapReader30")
		Reader.Load Connector.OutputStream
		' OBTENCION DE LOS DATOS



		Set xml=Reader.Body


		' call muestranodos(xml.childnodes)

		' tomo el idref de la persona
		Set nList = xml.selectNodes("//item")
		iz        = 0
		xgutop    = 0
		For Each n In nList
			id            = Replace(n.getAttribute("href") ,"#","")
			Set n2        = xml.selectSingleNode("//multiRef[@id='"&id&"']/name")
			idcomplemento = Replace(n2.getAttribute("href"),"#","")

			Set n3        = xml.selectSingleNode("//multiRef[@id='"&idcomplemento&"']")
			nombre        = n3.selectSingleNode("employeeName").text
			apellido      = n3.selectSingleNode("employeeLastname").text
			legajo        = n3.selectSingleNode("userId").text

			xxx           = CStr(UCase(apellido & " "  & nombre ))
            arrr(iz,1)    = replace( xxx & "" , "'"," ")
			'arrr(iz,1)    = xxx
			arrr(iz,0)    = legajo
			iz            = iz + 1
		Next

		xgutop    = iz

		If iz > 0 Then
			Call QSort(arrr,0,iz-1)

			For iz=0 To  xgutop

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & arrr(iz,0)
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & arrr(iz,1)

				xgusep                   = Chr(9)

			Next


		End  If

		Set Reader=Nothing
		Set Serializer=Nothing
		Set Connector=Nothing

		Case "P" ' Patagonia

		Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

		Dirc=Dirc & "CorEssDll\"

		Dirc=Dirc & "NSBTest_2_50_03.xml" ' busqueda  usuarios de un grupo
		f=Now()
		Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")
		xml4=leerArc(dirc)
		xml4=Replace(XML4,"XDRQCLUSR",Session("Usuario")&Rnd())
		xml4=Replace(XML4,"XFECHACOR",fechaX(F))
		xml4=Replace(XML4,"XHORACOR",horaMinSec(F))
		xml4=Replace(XML4,"XXXDATO",grupo) ' SE usa para buscar por texto

		xgusep=""
		session("SGU_UIGuserid")=""
		session("SGU_UIGdescri")=""
		xmlHttp2.open "POST", session("GetInfo"), False
		xmlHttp2.send (xml4)
		Set xmldoc=xmlHTTP2.responsexml



		Set z=xmldoc.selectSingleNode("//Evalls")

		Set nList = z.selectNodes("item")

		For Each n In nList

			If  N.hasChildNodes  Then 'primera ocurrencia tiene la info
				''response.write n.text
				xmldoc.LOADXML(N.text)
				Set n3 = xmldoc.selectNodes("//USUARIOS/Codigo")
				'  Set n4 = xmldoc.selectNodes("//USUARIOS/Mail")
				Set N5=  xmldoc.selectNodes("//USUARIOS/Nombre")

				For x= 0 To n3.length-1

					session("SGU_UIGuserid")=  session("SGU_UIGuserid") & xgusep & n3(x).text
					session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & n5(x).text

					xgusep  = Chr(9)

					'   mensaje=mensaje & "Codigo=" & n3(x).text &"<br>" & "MAIL=" & n5(x).text &"<br>" & "MAIL=" & n4(x).text  &"<br><br>"
				Next



				Exit For


			End If
		Next
		Set xmlHTTP2=Nothing
		Set xmldoc  =Nothing
		Set z	   =Nothing
		Case "CFA"

		getUsrIngroupADSV(grupo)


		Case "T"

		getUsrIngroupAD(grupo)

		Case "S"

		getUsrIngroupADSV(grupo)

		Case "BC"

		getUsrIngroupADSV(grupo)

		Case "GROI"

		getUsrIngroupBCRED(grupo)
		Case "BCOLU"
		getUsrIngroupADBCOLU(grupo)

		Case "LACAJA"

		getUsrIngroupADLaCaja(grupo)

		Case "RIVA"

		LegajosGrupo(grupo)

		Case "BANELCO"
		getUsrIngroupADBANELCO (GRUPO)


		Case "FINANSUR"
		getUsrIngroupAD_FINANSUR (GRUPO)

		Case "BPN"
		getUsrIngroupAD_BPN (GRUPO)

		Case "GP"
		getUsrIngroupAD_GP(GRUPO)
	End Select

End Sub
Sub getUsrInSect(XGU_SECT,SECSITE)
	Dim xml, root, i

	GTUcase = session("USRPRV")

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Select Case GTUcase

		Case "I"

		If IsNumeric(XGU_SECT) Then
			WXGU_SECT = XGU_SECT
		Else
			WXGU_SECT = 1
		End If

		Set xcn = ofv.conectar(ofv.strconn2)

		' xStrSql = "Select ID,descripcion  From USUARIOS WHERE SECTOR = 0" & WXGU_SECT  &" AND SITE2 = " & SECSITE & " order by nivel "
		xStrSql = "Select ID,descripcion  From USUARIOS WHERE SECTOR = 0" & WXGU_SECT  & " order by nivel "

		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then

			Do While Not XRS.EOF

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & xrs(1)

				xgusep = Chr(9)

				XRS.MOVENEXT
			Loop

		Else

		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)

		Case "Es"

		If IsNumeric(XGU_SECT) Then
			WXGU_SECT = XGU_SECT
		Else
			WXGU_SECT = 1
		End If

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select ID,descripcion  From USUARIOS WHERE SECTOR = 0" & WXGU_SECT & " order by descripcion "
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then

			Do While Not XRS.EOF

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & xrs(1)

				xgusep = Chr(9)

				XRS.MOVENEXT
			Loop

		Else

		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)

		Case "T"
		getUserInSectAd(XGU_SECT)

		Case "BC"
		GetUserInsectBC(XGU_SECT)

		Case "RIVA"
		LegajosSector(XGU_SECT)
		Case "BANELCO"
		getUserInSectAdBANELCO(XGU_SECT)

		Case "FINANSUR"
		getUserInSectAd_FINANSUR(XGU_SECT)
		Case "BPN"
		getUserInSectAd_BPN(XGU_SECT)

		Case "GP"
		getUserInSectAd_GP(XGU_SECT)

	End Select

End Sub
Sub getUsrInTEXT(XGU_TEXT)
	Dim xml,root,i

	GTUcase = session("USRPRV")

	xgusep                   = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Select Case GTUcase

		Case "I"

		WXGU_TEXT = Replace(XGU_TEXT & "","*","%")

		Set xcn = ofv.conectar(ofv.strconn2)
		If Session("EMPMUN")  = "S" Then
			xStrSql = "Select ID,descripcion  From USUARIOS WHERE site=" & session("SETIDEMPRE") & "  and DESCRIPCION LIKE '" & WXGU_TEXT & "' order by descripcion "
		Else
			xStrSql = "Select ID,descripcion  From USUARIOS WHERE DESCRIPCION LIKE '" & WXGU_TEXT & "' order by descripcion "
		End If
		'response.write xStrSql
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then

			Do While Not XRS.EOF

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & replace(xrs(1) & "" , "'"," ")

				xgusep                   = Chr(9)

				XRS.MOVENEXT
			Loop

		Else

		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"
		If session("APLMDCLI")="BGAL" Then

			WXGU_TEXT = Replace(XGU_TEXT & "","*","%")

			Set xcn = ofv.conectar(ofv.strconn2)

			xStrSql = "Select legajo, RTRIM(NOM_EMPLEADO) + ' ' + APE_EMPLEADO AS NOMBRE    From DIRGLOBAL_NOMINA WHERE  RTRIM(NOM_EMPLEADO) + ' ' + APE_EMPLEADO  LIKE '" & WXGU_TEXT & "' order by NOMBRE "
			'RESPONSE.WRITE xStrSql
			Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
			If Not xRS.EOF Then

				Do While Not XRS.EOF

					session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
					session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & replace(xrs(1) & "" , "'"," ")

					xgusep                   = Chr(9)

					XRS.MOVENEXT
				Loop

			Else

			End If

			Call ofv.cerrarconsulta(xrs)
			Call ofv.cerrarconn(xcn)


		End If
		Case "P"   ' patagonia
		Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

		Dirc=Dirc & "CorEssDll\"
		Dirc=Dirc & "NSBTest_2_50_05.xml" ' busqueda por texto
		f=Now()
		xgusep=""
		session("SGU_UIGuserid")=""
		session("SGU_UIGdescri")=""
		Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")
		xml4=leerArc(dirc)
		xml4=Replace(XML4,"XDRQCLUSR",Session("Usuario"))
		xml4=Replace(XML4,"XFECHACOR",fechaX(F))
		xml4=Replace(XML4,"XHORACOR",horaMinSec(F))

		xml4=Replace(XML4,"XXXDATO1","")  	' SE usa para buscar por id
		XGU_TEXT=Replace(XGU_TEXT,"*","")
		xml4=Replace(XML4,"XXXDATO",XGU_TEXT) ' SE usa para buscar por texto


		'exit sub
		xmlHttp2.open "POST", session("GetInfo"), False
		xmlHttp2.send (xml4)
		Set xmldoc=xmlHTTP2.responsexml
		'  response.write xmlHTTP2.responsetext
		Set z=xmldoc.selectSingleNode("//Evalls")

		Set nList = z.selectNodes("item")

		For Each n In nList

			If  N.hasChildNodes  Then 'primera ocurrencia tiene la info
				'response.write n.text
				xmldoc.LOADXML(N.text)

				Set nListx = XMLDOC.selectNodes("DATOS/USUARIO")
				For Each nx In nListx
					Set z1=nX.selectSingleNode("Codigo")
					Set z2=nX.selectSingleNode("Nombre")

					session("SGU_UIGuserid")= session("SGU_UIGuserid") & xgusep & Z1.text
					session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep  & Z2.text

					xgusep  = Chr(9)
				Next

				Exit For


			End If
		Next

		Set xmlHTTP2=Nothing
		Set xmldoc  =Nothing
		Set z	   =Nothing
		Case "CFA"

		WXGU_TEXT=Replace(XGU_TEXT & "","*","")

		WXGU_TEXT=WXGU_TEXT & "*"
		If WXGU_TEXT <> "*" Then
			getUsrInTEXTAD(WXGU_TEXT)
		End If

		Case "T"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD(XGU_TEXT)


		Case "LACAJA"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD(XGU_TEXT)

		Case "BANELCO"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTADBANELCO(XGU_TEXT)

		Case "FINANSUR"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD_FINANSUR(XGU_TEXT)
		Case "BPN"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD_BPN(XGU_TEXT)

		Case "GP"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD_GP(XGU_TEXT)

		Case "NOVABC"
		WXGU_TEXT=Replace(XGU_TEXT &"","*","") ' & "*"
		getUsrInTEXTBCEX( WXGU_TEXT)


	End Select

End Sub


Sub getUsrInTEXT2(XGU_TEXT)
	Dim xml,root,i

	GTUcase = session("USRPRV")

	xgusep                   = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Select Case GTUcase



		Case "I"
		WXGU_TEXT = Replace(XGU_TEXT & "","*","%")

		Set xcn = ofv.conectar(ofv.strconn2)

		'xStrSql = "Select ID,descripcion  From USUARIOS WHERE DESCRIPCION LIKE '" & WXGU_TEXT & "' order by descripcion "
		If Session("EMPMUN")  = "S" Then
			xStrSql = "Select ID,descripcion  From USUARIOS WHERE site=" & session("SETIDEMPRE") & "  and DESCRIPCION LIKE '" & WXGU_TEXT & "' order by descripcion "
		Else
			xStrSql = "Select ID,descripcion  From USUARIOS WHERE DESCRIPCION LIKE '" & WXGU_TEXT & "' order by descripcion "
		End If
		' response.write xStrSql
		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then

			Do While Not XRS.EOF

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & xrs(1)

				xgusep                   = Chr(9)

				XRS.MOVENEXT
			Loop

		Else

		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)


		Case "E"
		If session("APLMDCLI")="BGAL" Then

			WXGU_TEXT = Replace(XGU_TEXT & "","*","%")

			Set xcn = ofv.conectar(ofv.strconn2)

			xStrSql = "Select legajo, RTRIM(NOM_EMPLEADO) + ' ' + APE_EMPLEADO AS NOMBRE    From DIRGLOBAL_NOMINA WHERE  RTRIM(NOM_EMPLEADO) + ' ' + APE_EMPLEADO  LIKE '" & WXGU_TEXT & "' order by NOMBRE "
			'RESPONSE.WRITE xStrSql
			Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
			If Not xRS.EOF Then

				Do While Not XRS.EOF

					session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
					session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & replace(xrs(1) & "" , "'"," ")


					xgusep                   = Chr(9)

					XRS.MOVENEXT
				Loop

			Else

			End If

			Call ofv.cerrarconsulta(xrs)
			Call ofv.cerrarconn(xcn)


		End If

		Case "CFA"

		WXGU_TEXT=Replace(XGU_TEXT & "","*","")

		WXGU_TEXT=WXGU_TEXT & "*"
		If WXGU_TEXT <> "*" Then
			getUsrInTEXTAD(WXGU_TEXT)
		End If

		Case "T"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD(XGU_TEXT)


		Case "LACAJA"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD(XGU_TEXT)

		Case "BANELCO"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTADBANELCO(XGU_TEXT)

		Case "FINANSUR"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD_FINANSUR(XGU_TEXT)
		Case "BPN"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD_BPN(XGU_TEXT)

		Case "GP"
		WXGU_TEXT=WXGU_TEXT & "*"
		getUsrInTEXTAD_GP(XGU_TEXT)

		Case "NOVABC"
		WXGU_TEXT=Replace(XGU_TEXT &"","*","") ' & "*"
		getUsrInTEXTBCEX( WXGU_TEXT)

		Case "BC"
		WXGU_TEXT=Replace(XGU_TEXT &"","*","") ' & "*"
		getUsrInTEXTBCEX( WXGU_TEXT)


	End Select

End Sub

Sub QSort(aData, iLb, iUb)
	Dim lbStack(1000)
	Dim ubStack(1000)
	Dim sp              ' stack pointer
	Dim iLbx            ' current lower-bound
	Dim iUbx            ' current upper-bound
	Dim m
	Dim p               ' index to pivot
	Dim i
	Dim j
	Dim t               ' temp used for exchanges

	' Init stacks
	lbStack(0) = iLb
	ubStack(0) = iUb
	sp = 0

	Do While sp >= 0
		iLbx = lbStack(sp)
		iUbx = ubStack(sp)

		Do While (iLbx < iUbx)

			' select pivot and exchange with 1st element
			p = iLbx + (iUbx - iLbx) \ 2

			' exchange iLbx, p
			t = aData(iLbx,1)
			aData(iLbx,1) = aData(p,1)
			aData(p,1) = t


			t = aData(iLbx,0)
			aData(iLbx,0) = aData(p,0)
			aData(p,0) = t


			' partition into two segments
			i = iLbx + 1
			j = iUbx
			'RESPONSE.WRITE aData(iLbx,1) & "--" & aData(i,1) & "<BR>"
			Do
				Do While i < j
					If aData(iLbx,1) <= aData(i,1) Then Exit Do
					i = i + 1
				Loop

				Do While j >= i
					If aData(j,1) <= aData(iLbx,1) Then Exit Do
					j = j - 1
				Loop

				If i >= j Then Exit Do

				' exchange i, j
				t = aData(i,1)
				aData(i,1) = aData(j,1)
				aData(j,1) = t
				t = aData(i,0)
				aData(i,0) = aData(j,0)
				aData(j,0) = t

				j = j - 1
				i = i + 1
			Loop

			' pivot belongs in aData[j]
			' exchange iLbx, j
			t = aData(iLbx,1)
			aData(iLbx,1) = aData(j,1)
			aData(j,1) = t


			t = aData(iLbx,0)
			aData(iLbx,0) = aData(j,0)
			aData(j,0) = t


			m = j

			' keep processing smallest segment, and stack largest
			If m - iLbx <= iUbx - m Then
				If m + 1 < iUbx Then
					lbStack(sp) = m + 1
					ubStack(sp) = iUbx
					sp = sp + 1
				End If
				iUbx = m - 1
			Else
				If m - 1 > iLbx Then
					lbStack(sp) = iLbx
					ubStack(sp) = m - 1
					sp = sp + 1
				End If
				iLbx = m + 1
			End If
		Loop
		sp = sp - 1
	Loop
End Sub


'-----------------------------------------------------------------

Function webvalidabt()
	USUARIO=REQUEST.FORM("usuar")
	CLAVE=REQUEST.FORM("clave")
	webvalidabt=checkUsr(USUARIO,CLAVE)

End Function

Function  checkUsr(usuario,clave)

	'url="http://lx10s99e1.bancopatagonia.net.ar/bantotal/servlet/com.dlya.bantotal.ax9996c1"

	'url="http://nsbt.bancopatagonia.net.ar/bantotal/servlet/com.dlya.bantotal.ax9996c1"

	'url="http://desansbt.bancopatagonia.net.ar/bantotal/servlet/com.dlya.bantotal.ax9996c1"


	f=Now()


	Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

	Dirc=Dirc & "CorEssDll\test.xml"

	Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")
	xml4=leerArc(dirc)

	xml4=Replace(XML4,"XFECHACOR",fechaX(F))
	xml4=Replace(XML4,"XHORACOR",horaMinSec(F))
	xml4=Replace(XML4,"XUSUARIOCOR",USUARIO)
	xml4=Replace(XML4,"XPASSWORDCOR",CLAVE)

	xmlHttp2.open "POST", session("GetInfo"), False
	xmlHttp2.send (xml4 )

	Set xml3 =xmlHTTP2.responsexml

	checkUsr=validarUsu(xml3,usuario)

	Set xmlHTTP2=Nothing

End Function

Function leerArc(archivo)

	'response.write archivo
	Set FS = Server.CreateObject("Scripting.FileSystemObject")
	Set Fichero = FS.OpenTextFile(archivo, 1)

	leerArc=Fichero.readall

	Fichero.Close
	Set Fichero = Nothing

	Set FS = Nothing

End Function









Function validarUsu(xmldoc,usuario)

	Set z=xmldoc.selectSingleNode("//Evalls")

	Set nList = z.selectNodes("item")

	For Each n In nList

		If  N.hasChildNodes  Then 'primera ocurrencia tiene la info
			'response.write n.text
			xmldoc.LOADXML(n.text)
			Set z1=xmldoc.selectSingleNode("//MENSAJE")
			MENSAJE= z1.text

			Set z1=xmldoc.selectSingleNode("//SUCURSAL")
			SUCURSAL=z1.text

			Set z1=xmldoc.selectSingleNode("//NIVEL")
			nivel=z1.text

			Set z1=xmldoc.selectSingleNode("//NOMBRE")
			nombre=z1.text

			' Set z1=xmldoc.selectSingleNode("//PERFILGRUPO/CODIGO")
			' perfilgrupo=z1.text

			Set z2=xmldoc.selectSingleNode("//PERFILGRUPO")
			perfilgrupo=""
			Set nList2 = z2.selectNodes("Codigo")
			For Each nn In nlist2
				'Set z3=nn.selectSingleNode("//CODIGO")
				perfilgrupo=perfilgrupo & nn.text & Chr(9)
			Next


			Set z1=xmldoc.selectSingleNode("//EMAIL")
			EMAIL=z1.text

			Exit For


		End If
	Next




	If CStr( mensaje & "") =""  Or Len(mensaje & "") Then
		validarUsu="OK"
		session("mensajeweb")=""
		'  session("SGU_idcostcenter")= "0"
		' validarUsu= validarUsu & "SUCURSAL=" & SUCURSAL & "<br>"
		' validarUsu= validarUsu & "nivel=" & nivel & "<br>"
		' validarUsu= validarUsu & "nombre=" & nombre & "<br>"
		' validarUsu= validarUsu & "perfilgrupo=" & perfilgrupo & "<br>"
		' validarUsu= validarUsu & "EMAIL=" & EMAIL & "<br>"
		Session("Usuario")   = usuario
		Session("UsrOpcion") = nombre 'session("SGU_nombre")
		Session("SkrDefT")   = "0"
		Session("SkrDefE")   = "0"
		Session("SkrDefI")   = "0"
		Session("UsrPmin")   = "0"
		Session("UsrPmax")   = "0"
		Session("UsrHist")   = "0"
		Session("SegUsuario") = nombre 'session("SGU_nombre")
		Session("SegNivel")  = 8
		Session("SegGrp")    = "DEF"
		Session("SegPerfil") = usuario 'Session("Usuario")
		Session("UsuarNT")   = usuario 'Session("Usuario")

		Session("UsrSector") = "" 'session("SGU_idcostcenter")
		Session("UsrSecDsc") = "" 'session("SGU_descrisector")
		Session("UsrPuesto") = "" 'session("SGU_codigopuesto")
		Session("UsrPstDsc") = "" 'session("SGU_descripuesto")
		Session("UsrMail")   = EMAIL 'session("SGU_mail")



		Session("SegGrpOpt")    = "S"

		Session("SegGrpLst")    = perfilgrupo 'session("SGU_grupos")

	Else
		validarUsu=""
		session("mensajeweb")=mensaje
	End If



End Function

Function horaMinSec(f)

	hora=Hour(f)
	min = Minute(f)
	sec = Second(f)


	If hora < 10 Then hora="0" & hora
	If min < 10 Then min="0" & min
	If sec < 10 Then sec="0" & sec


	horaMinSec=hora & ":" & min & ":"  &sec
End Function


Function fechaX(f)

	mes=Month(f)
	dia=Day(f)

	If mes < 10 Then mes= "0" & mes
	If dia < 10 Then dia= "0" & dia

	fechax=Year(f)&"-" & mes & "-" & dia
End Function

'-------------------------------------------------------------------

Function getInfoUserExBp(Legajo)
	data = ""

	GTUcase = session("USRPRV")

	Select Case GTUcase

		Case "I"

		Set xcn = ofv.conectar(ofv.strconn2)

		xStrSql = "Select a.ID,a.descripcion,a.email,a.sector, a.puesto,a.legajo,a.site ,b.descri"
		xStrSql= xStrSql & " From USUARIOS a "
		xStrSql= xStrSql & " left outer join sites b on  (a.site=b.objid) "
		xStrSql= xStrSql &  " WHERE a.ID = '" & Legajo & "' "

		Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
		If Not xRS.EOF Then
			session("SGU_userid")      = xrs(0)
			session("SGU_nombre")      = xrs(1)
			session("SGU_mail")	   = xrs(2) & ""
			session("SGU_descrisector")= "MySect"
			session("SGU_idcostcenter")= xrs(3)
			session("SGU_descripuesto")= "MyPst"
			session("SGU_codigopuesto")= xrs(4)
			session("SGU_legajo")= xrs(5)
			session("SGU_site")= xrs(6)
			session("SGU_siteDescri")= xrs(7)
		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"
			session("SGU_legajo") = "NO"
			session("SGU_site") = "NO"
			session("SGU_siteDescri")="NO"
		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcn)





	End Select

End Function

Function webvalidaRivaex()
	USUARIO=REQUEST.FORM("usuar")
	CLAVE=REQUEST.FORM("clave")

	webvalidaRivaex=LOGUIN(USUARIO,CLAVE)
	session("mensajeweb")=""
	If webvalidaRivaex="" Then
		webvalidaRivaex="OK"
		Session("Usuario")   = usuario
		Call InfousrLOGUIN(USUARIO)
	Else
		session("mensajeweb")=webvalidaRivaex
	End If

End Function


Sub getGruposMail(XGU_TEXT)
	Dim xml,root,i

	GTUcase = session("USRPRV")

	xgusep                   = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Select Case GTUcase


		Case "BPN"
		WXGU_TEXT=WXGU_TEXT & "*"
		getGruposMail_BPN(XGU_TEXT)

	End Select

End Sub


Function getMailDelGrupo(XGU_TEXT)
	Dim xml,root,i

	GTUcase = session("USRPRV")

	xgusep                   = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	getMailDelGrupo=""
	Select Case GTUcase


		Case "BPN"
		WXGU_TEXT=XGU_TEXT '& "*"
		getMailDelGrupo=getMailDelgrupo_BPN(WXGU_TEXT)

	End Select

End Function


Function BinaryToString1(Binaryse, charset)
	'http://www.pstruh.cz/tips/detpg_binarytostring.htm

	Const adTypeText = 2
	Const adTypeBinary = 1
	'response.write Binaryse
	'Create Stream object
	Dim BinaryStream 'As New Stream
	If IsArray(Binaryse) Then
		Set BinaryStream = server.CreateObject("ADODB.Stream")

		'Specify stream type - we want To save text/string data.
		BinaryStream.Type = adTypeBinary

		'Open the stream And write text/string data To the object
		BinaryStream.open
		BinaryStream.write Binaryse

		'Change stream type To binary
		BinaryStream.position = 0
		BinaryStream.Type = adTypeText

		'Specify charset For the source text (unicode) data.
		BinaryStream.charset = charset
		'response.Write Charset

		'Open the stream And get binary data from the object
		BinaryToString1 = BinaryStream.ReadText
		Set BinaryStream=Nothing
	Else
		BinaryToString1=""
	End If
End Function


Sub Listarc(ByRef Directorio,ByRef resulta,wd)

	For Each Subdirectorio In Directorio.SubFolders
		'List1.AddItem Directorio.Path & "\" & Subdirectorio.Name
		XPATH=Directorio.Path
		X=InStr(XPATH ,wd)
		XPATH=Mid(XPATH,x,Len(XPATH))
		resulta=resulta  & XPATH  & "\" & Subdirectorio.Name & "<BR>"
		Listarc Subdirectorio, resulta,wd
	Next
End Sub

%>
