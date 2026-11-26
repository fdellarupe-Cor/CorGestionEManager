<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	Response.expires = 0
	'response.write request.querystring
	Server.scripttimeout = 600
	Response.buffer = True
	smenu = Request.querystring("Menu")


	CantRegMostrar = 12
	CantRegMover   = CDbl(Request.querystring("CantRegMover"))
	OrdenarPor     = Request.querystring("OrdenarPor")
	sNombre        = request.querystring("Nombre")

	sNombreV       = request.querystring("Nombre")
	smenuV         = Request.querystring("Menu")

	If  Len(Trim(OrdenarPor)) = 0 Then
		OrdenarPor  = "DESCRI"
	End If

	OrdenarPor      = UCase(OrdenarPor)&""
	sOrdenx         = Trim(OrdenarPor)
	' response.write OrdenarPor
	'response.write request.querystring & "<br>" & request.form & "<br>"
	Select Case OrdenarPor
		Case ""
		OrdenarPor = " order by complementos.descri"

		Case "DESCRI"
		OrdenarPor = " order by complementos.descri"
		Case Else
		OrdenarPor = " order by " & OrdenarPor
	End  Select



	smenu = UCase(smenu)

	If smenu = "PORTEXTO" Then

		If Request.Form.Count > 0 Then
			sIDv = Request.Form("buscar")

			sID = Replace(Request.Form("buscar"), "'", "")
			sID = Replace(sID, Chr(34), "")
		Else
			sIDv = Request.querystring("buscar")
			sID = Request.querystring("buscar")
			sID = OFV.convertircar(sID, "_", " ")
		End If

	ElseIf smenu = "PORIDENTIFICACION" Then

		If Request.Form.Count > 0 Then

			sID = Request.Form("buscar")
			sIDv = Request.Form("buscar")
		Else
			sID = Request.querystring("buscar")
			sIDv = Request.querystring("buscar")
		End If

	Else

		If Request.Form.Count > 0 Then

			sID = Request.Form("lista")
			sIDv = Request.Form("lista")
		Else
			sID = Request.querystring("buscar")
			sIDv = Request.querystring("buscar")
		End If

	End If

	If sID = "" Then
		sIDv = Request.querystring("buscar")
		sID = Request.querystring("buscar")
		sID = OFV.convertircar(sID, "_", " ")
	End If


	sID    = OFV.convertircar(sID, "_", " ")
	Set cn = OFV.conectar(OFV.strconn2)


	'Response.write smenu & "<br>"
	Select Case UCase(smenu)
		Case "PORTIPO"

		If session("EMPMUN") = "S" Then
			strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 "
			strsql = strsql & " From Complementos,TIPOCOMP Where TIPOCOMP=TIPO and compUN=" & session("UN") & " AND tipocomp='" & sID & "' " & OrdenarPor
		Else
			strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 "
			strsql = strsql & " From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND tipocomp='" & sID & "' " & OrdenarPor
		End If
		'    Response.write strsql & "<br>"
		sTitulo = " Tipo Componente: " & sID

		Case "PORTEXTO"
		sTitulo = "Texto: " & sID

		If session("EMPMUN") = "S" Then
			strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 From Complementos,TIPOCOMP "
			strsql = strsql & " Where compUN=" & session("UN") & " and TIPOCOMP=TIPO AND Complementos.DESCRI  like '%" & sID & "%' " & OrdenarPor
		Else
			strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND Complementos.DESCRI  like '%" & sID & "%' " & OrdenarPor
		End If

		'  Response.write strsql & "<br>"
		sTitulo = "Texto: " & sID
		Case "PORIDENTIFICACION"

		If session("EMPMUN") = "S" Then
			strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 From Complementos,TIPOCOMP "
			strsql = strsql & " Where compUN=" & session("UN") & " and TIPOCOMP=TIPO AND codiden ='" & sID & "'  " & OrdenarPor
		Else
			strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 From Complementos,TIPOCOMP "
			strsql = strsql & " Where TIPOCOMP=TIPO AND codiden ='" & sID & "'  " & OrdenarPor
		End If
		'   Response.write strsql & "<br>"
		sTitulo = "IDENTIFICACION: " & sID

		Case "PORSISTEMA"

		If session("EMPMUN") = "S" Then
			strsql = "select Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 from Compprty,complementos,TIPOCOMP"
			strsql = strsql & " where complementos.compUN=" & session("UN") & " and sistorigen='" & sID & "' "
			strsql = strsql & " and Compprty.codcomp=complementos.codcomp "
			strsql = strsql & " and Compprty.tipocomp=complementos.tipocomp "
			strsql = strsql & " and complementos.tipocomp=TIPOCOMP.TIPO  " & OrdenarPor
		Else
			strsql = "select Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 from Compprty,complementos,TIPOCOMP"
			strsql = strsql & " where sistorigen='" & sID & "' "
			strsql = strsql & " and Compprty.codcomp=complementos.codcomp "
			strsql = strsql & " and Compprty.tipocomp=complementos.tipocomp "
			strsql = strsql & " and complementos.tipocomp=TIPOCOMP.TIPO  " & OrdenarPor
		End If
		' Response.write strsql & "<br>"
		sTitulo = "Sistema: " & sID

	End Select

	'Response.write strsql & "<br>"

	Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

	If Not rs.eof Then
		If UCase(sMenu) = "PORTIPO" Then
			sTitulo = "Tipo Componente: " & rs("descri2")
		End If
	End If

	sHTML = OFV.FormHeader(sTitulo)
	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	xs = xs & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," '& Chr(34)
	xs = xs & "'menubar=no,toolbar=no,height='+h+',width='+w+',resizable=yes'" '& 'Chr(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)


	'Inicio Modificacion Mariano 25/09/2013

	'*****************************************
	'response.write request.querystring
	' response.write  "   esta es la opcion elegida: " & sNombren
	' response.write "</br>"
	' response.write "MENU: " & sMenu

	sAccion = "CWListadoConsultasComp.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&Nombre=" & sNombre & "&buscar=" & sID
	''   sAccion = "CWListadoConsultasComp.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre& "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")
	sInput   = ofv.GenerarInput("","Vista Web","submit","","10","bt","")

	' sInput   = sinput & ofv.GenerarInput("lista",sID,"hidden","","1","dt","")
	' sInput   = sinput & ofv.GenerarInput("lista2",sID,"hidden","","1","dt","")
	Formvista = ofv.GenForm("vistar","",sAccion,"","post"," style='display:none;' ",sInput,"si") & "Cambiar a Vista Web"

	Celda  = OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;'  ", "", "", "", "", "", sImagena & "&nbsp;" & sTitulo & "", "si")
	Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ","center onclick=""document.getElementById('vistar').submit();"" ","20%","","","",Formvista,"si")
	'   Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;' ","right","20%","","","","&nbsp;" & sTitOrd2 & "Orden:&nbsp;" & sTitOrd & "&nbsp;&nbsp;","si")



	Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



	tablaz = OFV.GenTabla("", "aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'", "", "100% ", "", "0", "", "1", "2", "", "", Fila, "si")



	sHTML = sHTML & Tablaz

	' Fin Modificacion Mariano 25/09/2013


	sHTML = sHTML & xs

	sID = OFV.convertircar(sID, " ", "_")
	Celda = OFV.GenCelda("", "ttr align=center ", "", "", "", "", "", "Tipo", "si")
	sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=Descri&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Descripción</a>"
	Celda = Celda & OFV.GenCelda("", "ttr  align=center ", "", "", "", "", "", sLink, "si")
	sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=descri2&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Tipo&nbsp;Componente</a>"
	Celda = Celda & OFV.GenCelda("", "ttr  align=center ", "", "", "", "", "", sLink, "si")
	sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=codiden&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Identificación</a>"
	Celda = Celda & OFV.GenCelda("", "ttr  align=center  ", "", "", "", "", "", sLink, "si")
	Fila = OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

	Response.write sHTML

	sHTML = ""

	Response.flush
	nLinea = 0
	PosReg = 0
	On Error Resume Next

	xTrColor = "#FFFFFF"


	'response.write strsql

	If Not rs.EOF Then rs.Move (CantRegMover)

	If rs.eof  Then

		'XFRASE = "<BR><BR>&nbsp;<B><font color=#993333 SIZE=3 FACE='corflaticon' >i</font></B>"
		XFRASE = "<BR><BR><i class='las la-info-circle' style=color:#993333;font-size:10pt; ></i>"
		XFRASE = XFRASE &  "&nbsp;NO SE ENCONTRARON DOCUMENTOS PARA LA BUSQUEDA SOLICITADA<BR><BR>"

		Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", XFRASE, "si")


		Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")
	End If

	Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

		'sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"
		sImagen = "<i class='las la-plus';></i>"

		sHref2       = "EditarImagen2.asp?ID=" & rs(0)
		sHref        = "javascript:void(0);"

		sLink = OFV.GenLink("", "nt style=opacity:1; onclick=""AbrirCons('" & sHref2 & "');return false;"" ", sHref, "_blank", "", "", sImagen, "si")

		Celda = OFV.GenCelda("", "ut", "center", "2%", "", "", "", sLink, "si")


		sLink = OFV.GenLink("", "nt onclick=""AbrirCons('" & sHref2 & "');return false;"" ", sHref, "_blank", "", "", rs(3), "si")

		Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", sLink, "si")
		Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", rs(6), "si")
		Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", rs(1), "si")

		Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

		If xTrColor = "#FFFFFF" Then
			xTrColor = "#F9F9F9"
		Else
			xTrColor = "#FFFFFF"
		End If

		PosReg = PosReg + 1
		rs.movenext
		nLinea = nLinea + 1

		If nLinea > 20 Then
			nLinea = 0
			Response.flush
		End If

	Loop



	sHTML = sHTML & OFV.GenTabla("", "aaa style='border-radius: 0px 0px 10px 10px;'", "", "95% ", "", "0", "", "-1", "0", "", "", Fila, "si")

	sAccion = "SubconsultaComponentes.asp?Menu=" & sMenu & "&Nombre=" & sNombre
	Celda   = ofv.BotonVolver(sAccion)

	Addqry = "&Menu=" & smenuv & "&OrdenarPor=" & sOrdenx

	Addqry = Addqry & "&buscar=" & sIDv & "&nombre=" & snombrev

	sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)


	sHTML = sHTML  & "</body></html>"

	Call OFV.cerrarconsulta(rs)
	Call OFV.cerrarconn(cn)

	response.write sHTML
	response.flush

	sHTML = ""
End If



If request.querystring("scal")="S" Then
	strsql=session("sqlcons")
Else
	session("sqlcons")=strsql
End If

'' response.write strsql

response.write sHTML
%>