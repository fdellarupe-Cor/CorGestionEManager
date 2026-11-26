<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else

	'set oPrtComp = server.createobject("CorGeneratorComp.Reportes")
	'Set oPrtComp.Request = Request
	'Set oPrtComp.Server = Server
	'Set oPrtComp.Response = Response
	'Set oPrtComp.Session = Session
	'Set oPrtComp.Ofv = Ofv
	''
	'sbloque1= oPrtComp.generar()
	'
	'
	'set oPrt = server.createobject("CorPrintersp.Reportes")
	'
	'Set oPrt.Request = Request
	'Set oPrt.Server = Server
	'Set oPrt.Response = Response
	'Set oPrt.Session = Session
	'Set oPrt.Ofv = Ofv

	'if ofv.explorador = "MSIE" then
	'   oPrt.AltoPagina = 185
	'else
	'   oPrt.AltoPagina = 130
	'end if

	id = Request.querystring("ID")
	Set cn = ofv.conectar(ofv.strconn2)
	strsql="Select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND codcomp=" & id
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	sRepHead = RepHeader2(rs("descri2") & "&nbsp;:&nbsp;"& rs("codiden") & "&nbsp;-&nbsp; " & rs("descri"), "600")
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	sPath = Server.mappath(Request.ServerVariables("SCRIPT_NAME"))
	sPath2 = Mid(sPath, 1, InStrRev(sPath, "\", -1))
	sHead = OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")

	Celda = OFV.GenCelda("", "vt", "right", "100%", "", "", "", "PAGINA&nbsp;&&XX de &&XY", "si")
	sHead = sHead & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


	Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
	sHead = sHead & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenFTabla()


	'sError = oPrt.renderizar(cstr(sbloque1),cstr(shead),cstr(sRepHead),cstr(sPath2))
	'response.write oPrt.sPath2
	'if sError <> "" then response.write sError

	'end if



	Function Generar()

		If SESSION("URLPORT")  = "N" Then
			ruta= TipoSite("")  & request.servervariables("SERVER_NAME")  & Request.ServerVariables("SCRIPT_NAME")
		Else
			ruta= TipoSite("")  & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & Request.ServerVariables("SCRIPT_NAME")
		End If
		'ruta="http://" & Request.ServerVariables("server_NAME") & Request.ServerVariables("SCRIPT_NAME")
		'response.write ruta
		ruta = Mid(ruta, 1, InStrRev(ruta, "/", -1))

		Generar = "<html>"
		generar=generar & ofv.formstyler() & "<body>"
		Set cn = ofv.conectar(ofv.strconn2)

		Response.buffer = True
		Server.scripttimeout = 600
		id = Request.querystring("ID")



		'estructura
		strsql = "Select * From compestr Where codcomp = " & ID & " order by objid "
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		strsql = "Select * From compdet Where codcomp = " & ID & " order by codestr,datoformul,objid"
		Set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


		Fila = ""
		sestr = "no"
		Do While Not rs.eof
			sestr = "si"
			'Estructura
			Celda = ofv.GenCelda("","tu colspan=2","","","","","","Estructura: " & rs(3),"si")
			Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
			Celda = ofv.GenCelda("","tu","","","","","","Descripción","si")
			Celda = Celda & ofv.GenCelda("","tu","","","","","","Tipo","si")
			Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
			sdeta = "no"
			sind = 0
			'Detalle
			Do While Not rs2.eof
				If CStr(rs(0)) = CStr(rs2(3)) Then
					sdeta = "si"
					sind = sind + 1
					Celda = ofv.GenCelda("","dt","","70%","","","",CStr(sind) & " - " & rs2(4),"si")
					Celda = Celda & ofv.GenCelda("","dt","","30%","","","",rs2(5),"si")
					Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
					rs2.movenext
				ElseIf CStr(rs(0)) > CStr(rs2(3)) Then
					rs2.movenext
				Else
					Exit Do
				End If
			Loop

			If sdeta = "no" Then
				Celda = ofv.GenCelda("","dt","","70%","","","","Sin Datos","si")
				Celda = Celda & ofv.GenCelda("","dt","","30%","","","","&nbsp;","si")
				Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
			End If

			rs.movenext
		Loop

		If sestr = "no" Then
			sestr = "<center><font color=#336699><h4><p align=center>Sin Estructura disponible"
			sestr = sestr & "</p></h4></font></center>"
			sestr=""
			Celda = ofv.GenCelda("","xx colspan=2","","70%","","","",sestr,"si")
			Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
		End If

		sHTML = ofv.FormBody2 & ofv.GenTabla("","","","100% ","","0","0","0","0","","",Fila,"si")



		Call ofv.cerrarconsulta(rs)
		Call ofv.cerrarconsulta(rs2)


		'imagenes
		strsql = "Select * From Compimag Where codcomp = " & ID

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		i = 0
		'if rs.eof then
		'sHTML = sHTML & "No se cargaron datos"
		'end if

		ReDim Preserve aImagen(i)
		ReDim Preserve aDescri(i)
		ReDim Preserve aSecuen(i)
		Do Until rs.eof
			ReDim Preserve aImagen(i)
			ReDim Preserve aDescri(i)
			ReDim Preserve aSecuen(i)
			aImagen(i) = rs(3)
			aDescri(i) = rs(4)
			aSecuen(i) = rs(5)
			rs.movenext
			i = i + 1
		Loop

		'
		Call ofv.cerrarconsulta(rs)
		Call ofv.cerrarconn(cn)


		celda=""
		Celda =""' ofv.GenCelda("","","","","","","","&nbsp;","si")
		Fila = ""' ofv.GenRow("","","","","","","",Celda,"si")

		For i = LBound(aImagen) To UBound(aImagen)
			'celda =""' ofv.GenCelda("","","","","","","","&nbsp;","si")
			'  Fila = ""' ofv.GenRow("","","","","","","",Celda,"si")

			If aImagen(i) = "" Then
				sInput = ""
				sImagen = ""
			Else
				sInput = "<div>&nbsp;</div><Font class=at> Descripción: " & aDescri(i) & " Cuerpo: " & aSecuen(i) &"<br><br></font>"

				' Celda = ofv.GenCelda("","tu ","","","","","",sinput,"si")
				' Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
				If Session("EMPMUN") = "S" Then
					sImagen =  ofv.GenImage("","","","","Emp" & session("UN") & "/componentes/" & aImagen(i),"","") ' & "<div>&nbsp;</div>"
				Else
					sImagen =  ofv.GenImage("","","","","componentes/" & aImagen(i) ,"","") ' & "<div>&nbsp;</div>"

				End If

				' response.write ruta & "222"
			End If
			'Celda = ofv.GenCelda("","","","","","","", simagen,"si")
			'Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
			'Celda = ofv.GenCelda(""," ","","","","","", "<div>&nbsp;</div>","si")
			'Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
			's1HTML = sHTML & sInput & "</p></h4>" & sImagen
			sHTML = sHTML &  sInput  & Chr(13) & simagen   & Chr(13) 'ofv.GenTabla("","","","100%  ","","0","0","0","0","","",Fila,"si")

		Next
		'sHTML = sHTML & "</font>"
		'Celda = ofv.GenCelda(""," ","","","","","",shtml ,"si")
		'Fila =  ofv.GenRow("","","","","","","",Celda,"si")
		'sHTML = sHTML & ofv.GenTabla("","","","100%  ","","0","0","0","0","","",Fila,"si")

		Generar =  sHTML   & "</body></html>"



	End Function

	Public Function RepHeader2(STitulo, nAncho)

		Dim s, Fecha

		If nAncho = "" Then nAncho = "600"

		s = "<table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
		s = s & "<tr align=top>"
		If Session("CliLogo") <> "" Then s = s & "<td width=4% ><img src='" & Session("CliLogo") & "' border=0 ></td>"
		s = s & "<td class=ut align=left >&nbsp;" & Session("Cliente") & "</td></tr></table>"
		s = s & "<br><table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
		s = s & "<tr align=top><td class=vt align=left width=70% >" & STitulo & "</td>"
		Fecha = OFV.fechador()
		s = s & "<td class=ut align=right>" & Fecha & "</td></tr></table>"
		s = s & "<hr color=#000088 size=1 width=" & nAncho & " align=left>"

		RepHeader2 = s

	End Function
	Function cnumer(cdato )
		Dim ldato
		Dim kk

		ldato = Len(cdato)

		For kk = 1 To ldato
			If IsNumeric(Mid(cdato, kk, 1)) Then
				cnumer = cnumer & Mid(cdato, kk, 1)
			End If
		Next


	End Function

	response.write  Generar()
End If


%>