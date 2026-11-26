<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	If request.querystring("seg") <> "" Then
		Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
	End If


	XXXX = NEWEMPRE("NO")

	session("ASP")="SUBCONSULTA.ASP"

	sUN = request.querystring("EMPRE")

	If sUN = ""  And SESSION("UN") = "" Then

		sUN = ATREMP(0, 0)
		SESSION("UN") = sUN & ""
		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
	ElseIf sUN <> "" Then
		XXXX = NEWEMPRE(sUN)
		SESSION("UN")=sUN
	Else
		XXXX = NEWEMPRE(SESSION("UN"))
		sUN=SESSION("UN")
	End If

	'response.write SESSION("UN")
	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar

	vuelta = ofv.vueltaasp2("../")
	vuelta2 = ofv.vueltaasp2("")


	sHTML = ofv.FormHeader("Autorizantes Default")

	TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


	shtml = shtml & tablah & ""



	Set cn = ofv.conectar(ofv.strconn2)
	Set cn3 = ofv.conectar(ofv.strconn6)


	Fila  = ""
	Celda = ofv.gencelda("","ttt colspan=6","center","2%","","","","Lista de Autorizantes","si")
	Fila  = fila & ofv.GenRow("","","","","25","","",Celda,"si")

	Celda = ofv.gencelda("","ttr","","2%","","","","&nbsp;","si")
	Celda = Celda & ofv.gencelda("","ttr","","2%","","","","Orden","si")
	Celda = Celda & ofv.gencelda("","ttr","","60%","","","","Usuario","si")

	'''Nuevo -15012014-
	'CIRCUITO
	Celda = Celda & ofv.GenCelda("","ttr","","40%","","","","Autorizacion","si")

	'if session("EMPMUN") = "S" then
	'Celda = Celda & ofv.GenCelda("","ttr","","40%","","","","Empresa","si")
	'end if

	Fila = ofv.GenRow("","","","","","","",Celda,"si")
	Formulario = Fila

	celda = ""

	''
	'''Nuevo -15012014- Fin

	'Si es Multiempresa entonces mira el alcance
	'sino solo el codigo de objeto = 0

	IF Session("EMPMUN") = "S" THEN
		Strsql = "SELECT * FROM OBJwfldet  "
		Strsql = Strsql & " Where codobjeto=0 and alcance=" & session("UN")
		Strsql = Strsql & " order by orden,objid"
			ELSE
				Strsql = "SELECT * FROM OBJwfldet  "
				Strsql = Strsql & " Where codobjeto=0 "
				Strsql = Strsql & " order by orden,objid"
	END IF

	'response.write strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	''  Formulario = ""
	vinoSC = 0

	Do Until rs.eof

		sForm = "wfl" & rs(0)

		If bModificar Then
			sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objwfldet&volver="
			sHref = sHref  & vuelta
			sHref = sHref & "&Codobjeto=0&Tipobjeto=ALL&Eliminar=NSNC"
		End If
		If sHref <> "" Then
			Celda = ofv.BotonEliminar("Usuario",sHref)
		Else
			Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
		End If

		'Orden
		If bModificar Then
			sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"2","dtn","SI")
			sInput = ofv.gencelda("","ut","","1%","","","",sInput,"si")

		Else
			sinput = ofv.gencelda("","ut","","1%","","","",rs(3),"si")
		End If
		Celda = Celda & sinput

		sTabla = "Usuario"

		sBusca = rs(4)

		sGrpUsr = "S/D"
		'set cx = ofv.conectar(ofv.strconn0)

		Strsql = "SELECT * FROM " & sTabla & "s"
		Strsql = Strsql & " Where ID = '" & sBusca & "' "

		'set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
		'if not rx.eof then sGrpUsr = sTabla & ": " & rx(0) & " - " & rx(1)
		'call ofv.cerrarconsulta(rx)
		'call ofv.cerrarconn(cx)
		'response.writers(4)
		'xxx=getinfouser(sbusca)

		xxxx     = getInfoUserEX(sBusca,rs(14))
		sGrpUsr = sTabla & ": " & sBusca & " - " & session("SGU_nombre")

		sinput = ofv.gencelda("","ut ","","100%","","","",sGrpUsr,"si")

		Celda = Celda & sinput

		'' Nuevo 15012014
		SEGWHERE = " WHERE INI = 1 AND INICIO = 'A' AND ( (OBJTIPO = 'COR')) "
		StrSql1 = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
		' response.write strsql1
		scomKK = ""
		scomDD = ""
		''  response.write strsql & "<br>" & chr(13)
		Set rsniv = ofv.crearconsultaEx(StrSql1,cn3,1,parametros)
		Do Until rsniv.eof
			scomKK = scomKK & rsniv(0) & Chr(9)
			scomDD = scomDD & rsniv(1) & Chr(9)
			RSNIV.MOVENEXT
		Loop
		Call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2).value,scomKK,scomDD,"")

		'' sInput = ofv.GenerarCombo("Flform","Fltipo",Fltipo,Strsql,cn,0,1,"si")
		'sInput = ofv.GenerarCombo(sForm,rs(2).name,rs(2).value,StrSql1,cn3,0,1,"")
		Celda = Celda & ofv.gencelda("","VTI","","80%","","","",sInput,"si")


		'IF session("EMPMUN") = "S" then

		'    sImagenP = "&#0232;"
		'	sinput = "<a class='ft1o' href='AutDefaultDetEmp.asp?ID=" & rs(0) & "' "
		'    sinput = sinput  & " title='Seleccionar alcance' >" & sImagenP & "</a>"

		'    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

		'end if

		Fila = ofv.GenRow("","","","","","","",Celda,"si")

		sAccion = ""
		If bModificar Then
			sAccion = "asp/actualizar.asp?fname=" & rs(0)
			sAccion = sAccion & "&tabla=OBJwfldet&volver=" & vuelta
		End If
		'' Nuevo 15012014 Fin

		Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

		rs.movenext
	Loop



	sHTML = sHTML & "<center>"
	''''''''sHTML = sHTML & ofv.GenTabla("","aa  border=0","","70%","","0","0","1","1","","",Formulario & fila,"si")
	sHTML = sHTML & ofv.GenTabla("","aaa ","","70%","","0","0","1","1","","",Formulario,"si")

	sHTML = sHTML & "</center>"

	Call ofv.cerrarconsulta(rs)

	Call ofv.cerrarconn(cn)

	'Botonera
	Agregar = ""
	If bAgregar Then
		sAccion = "AgregarLista.asp?caller=AUT&Codobjeto=0&indice=NO"
		sAccion = sAccion & "&Tipobjeto=ALL&cod=ALL&volver=" & vuelta
		Agregar = ofv.BotonAgregar(sAccion)
	End If

	sAccion = "AutDefault.asp"
	Celda = ofv.BotonVolver(sAccion)

	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,Celda)


	sHTML = sHTML & "</body></html>"
End If

Function selecEmpre()

	sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

	If  Session("EMPMUN") = "S" Then

		Fila    = ""

		SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

		Fila    = ""
		Celda   = ""
		sHref   = ""

		sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
		Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
		Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
		Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

		Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

		Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





		shtml=shtml & "<script language=javascript>" & Chr(13)
		shtml=shtml & "function shwDem(){" & Chr(13)
		shtml=shtml & "document.all.DEM1.style.display='block';" & Chr(13)
		shtml=shtml & "}" & Chr(13)
		shtml=shtml & "</script>" & Chr(13)
		'----------------------------------------------
		sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "document.all.DEM1.style.display='none';" & Chr(34)
		sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
		TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


		TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
		TABLAX = TablaX & "<tr><td  "
		TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
		TABLAX = TablaX & "<tr><td colspan=3 "
		TABLAX = TablaX & " >&nbsp;</td></tr>"

		For IAEMP=1 To TAEMP

			'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

			If ATREMP(IAEMP,2) = 1 Then

				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
				Else
					TABLAX = TablaX & "<tr>"
				End If

				TABLAX = TablaX & "<td width=5 "
				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
				Else
					TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
				End If

				sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
				TABLAX = TablaX & "<td width=5 "
				TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

				TABLAX = TablaX & "<td class=vt  >"

				'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
				sHref = "AutDefaultDet.asp?EMPRE=" & ATREMP(IAEMP,0)
				'''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
				sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
				TABLAX = TablaX & sLink
				TABLAX = TablaX & "</td>"
				TABLAX = TablaX & "</tr>"

			End If

		Next

		TABLAX = TablaX & OFV.genFtabla()

		TABLA = Tabla & TABLAX

		TABLA = Tabla & "</DIV>"

		sHTML = sHTML & TABLA

	End If

End Function

response.write sHTML

%>