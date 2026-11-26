<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	retorno = request.querystring("retorno")
	Tema = request.querystring("Tema")
	CantRegMostrarx = 8
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	vuelta2 = ofv.vueltaasp("",CantRegMostrarx,Cantregmoverx)

	If tema = "" Then
		Tema = session("eltema")
	Else
		session("eltema") = CDbl(tema)
	End If
	If retorno = "" Then
		retorno = session("vuelta")
	Else
		session("vuelta") = CDbl(retorno)
	End If

	Set cn = ofv.conectar(ofv.strconn2)

	sHTML = ofv.FormHeader("SUBTEMAS")

	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar

	'Encabezados
	If session("UN") = 1 Then
		Fila = ofv.Encabezados("Descripcion","orden","Tipo","modo","redir.","seg.","Prop.")
	Else
		Fila = ofv.Encabezados("Descripcion","orden","Tipo","modo","redir.","seg.")
	End If

	Strsql = "Select tipobjeto,descri From Subtema"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	n = 0
	Do Until rs.Eof
		ReDim Preserve Perfiles(n)
		ReDim Preserve Perfilesx(n)
		Perfiles(n) = rs(0)
		Perfilesx(n) = rs(1)
		rs.Movenext
		n = n + 1
	Loop
	Call ofv.cerrarconsulta(rs)

	Strsql = "Select * From Prodserv where codtema = " & Tema & " Order By descri"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then rs.move(CantRegMoverx)
	posI=0
	Do While (Not rs.EOF) And (posi< CantRegMostrarx)
		Formulario = ""
		sForm = "form" & posi

		'Eliminar
		If bSuprimir Then
			sHref = "asp/EliminarSubtema.asp?ID=" & ofv.convertircar(rs(0)," ","_")
			sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC"
			Celda = ofv.BotonEliminar(rs(1),sHref)
		Else
			Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
		End If

		'Descripcion
		sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"150","dt","")
		Celda = Celda & ofv.GenCelda("","","","5%","","","",sInput,"si")

		'orden
		sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"5","dtn","")
		Celda = Celda & ofv.GenCelda("","","","5%","","","",sInput,"si")

		'Tipo
		sID = "document." & sForm
		sql1 = "Select tipobjeto,descri From Subtema where tipobjeto = '" & rs(4) & "' "
		Set rx = ofv.crearconsultaEx(sql1,cn,1,parametros)
		If Not rx.eof Then stipdes = rx(1)
		Call ofv.cerrarconsulta(rx)

		Celda = Celda & ofv.GenCelda("","cc","","5%","","","",stipdes,"si")


		'tipo
		Strsql = "Select codobjmaster From relaobj where codobjmaster = " & rs(0) & " and tipobjmaster = '" & rs(4) & "' "
		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsx.eof Then

			If rs("PSTIPO")="K" Then

				sInput = ofv.GenerarInput("",rs("PSTIPO"),"readonly","","20","dt","")
			Else
				sInput = ofv.GenerarInput("","Carpeta","readonly","","20","dt","")
			End If

		Else

			scomK = "C" & Chr(9) & "R"
			scomD = "Carpeta" & Chr(9) & "Redireccion"


			If bModificar And RS("PSTIPO") <> "K" Then
				sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),scomK,scomD," ")
			Else
				sInput = ofv.GenerarInput("",rs(7),"readonly","","10","dt","")
			End If

		End If

		Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


		If rs(7) = "R" Then

			'Descripcion
			sInput = ofv.GenerarInput(rs(8).name,rs(8),"text",sForm,"30","dt","")
			Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


		Else


			sinput = "----"
			Celda = Celda & ofv.gencelda("","ut","center","","1%","","",sInput,"si")


		End If




		If rs("segaceso") = "0" And rs("segactual") = "0" Then
			scl = "ft3"
		Else
			scl = "ft2"
		End If


		'sImagenP = "&#0232;"
		'sFuncAnc = " onmouseover=""this.className='" & scl & "o';"" onmouseout=""this.className='" & scl & "n';"" "
		sinput = "<a href='subtemaskr.asp?volver=" & vuelta2 & "&id=" & rs(0)
		sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
		sinput = sinput  & "' title='Permisos de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
		Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

		If session("UN") = 1 Then

			'sImagenP = "&#0232;"
			'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
			sinput = "<a href='ATRMTEMAST.asp?CantRegAMover=" & retorno & "&OPCION=" & rs(0)
			sinput = sinput  & "&T=S&NOMBRE=" & Replace(RS(1)," ","_")
			sinput = sinput  & "' title='Propiedades de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
			Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

		End If
		sAccion = ""
		If bModificar Then
			sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
			sAccion = sAccion & "&tabla=Prodserv&volver=" & vuelta
		End If

		Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")

		Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

		posi = posi+1
		rs.Movenext
	Loop

  ftab = ofv.GenTabla("","","","","","0","","0","0","","",Fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

	Agregar = ""
	If bAgregar Then

		sAccion = "AgregarSUBTEMAMAIN.asp?Regisvuelta=" & CantRegMoverx & "&Tema=" & Tema & "&Opcion=si&tabla=PRODSERV"
		Agregar = ofv.BotonAgregar(sAccion)
	End If


	sAccion = "TEMASsubtemas.asp?CantRegMover=" & retorno
	volverx = ofv.BotonVolver(sAccion)

	str1 = "Select descri From Prodserv where codtema = " & Tema & " Order By descri"    ' orden pagina
	str2 = "Select descri From Prodserv where codtema = " & Tema & " Order By descri"    ' orden alfabetico
	Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

	sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,volverx,"")

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)

	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>