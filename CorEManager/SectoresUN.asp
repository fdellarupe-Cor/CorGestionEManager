<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	CantRegMostrarx = 10
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

	sUN = request.querystring("UN")
	'response.write "CantRegMoverx: " & sUN & "  "  & session("sun")

	XXXX = NEWEMPRE("NO")

	session("ASP")="SUBCONSULTA.ASP"

	sUN = request.querystring("EMPRE")
	'response.write "sEMPRE: " & sUN
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
    if session("sun")="" then  session("sun") = sUN
	sHTML = ofv.FormHeader(session("FormN"))

	If SESSION("APLCHK") <> "S" Then xxxx = APLCHECK()

    ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar

	Set cn = ofv.conectar(ofv.strconn2)


	TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


	shtml = shtml & tablah '& "<br>"


	sUN = session("UN")



	'Encabezados

	If  session("APLMDLOR") <> "BE"  Then
		If  session("EMPMUN") = "S" Then
		if session("UN") = 1 then
    		Formulario = ofv.Encabezados("Id","Descripcion del Sector","Propiedades","ESTADO")
		Else
			Formulario = ofv.Encabezados("Id","Descripcion del Sector","Propiedades","Mapeo","ESTADO")
		End If
		else
		Formulario = ofv.Encabezados("Id","Descripcion del Sector","Propiedades","ESTADO")
		end if
	Else
		   Formulario = ofv.Encabezados("Id","CODIGO","Descripcion del Sector","Propiedades")
	End If

	If session("APLMDLOR") <> "BE"  Then
		Sordenar="	Order By DESCRI "
	Else
		Sordenar = " Order By CODINTER  "
	End If

	'If session("UN") <> 1 or Session("EMPMUN") = "N" Then
		Strsql = "Select * From ECO where ecoun = " & session("UN") & " order by descri "
	'Else
		'Strsql = "Select a.*,b.cliente From ECO a inner join clienteplan b on (a.ecoUN=b.ID) order by b.cliente"
	'	Strsql = "Select a.*,b.cliente From ECO a inner join clienteplan b on (a.ecoUN=b.ID) order by b.cliente"
	'End If

	'response.write strsql

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	If Not rs.EOF Then   rs.move(CantRegMoverx)

	posI = 0

	Do While (Not rs.EOF) And (posi < CantRegMostrarx)
		sForm = "form" & posi

		'Eliminar
		If bSuprimir Then
			sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
			sHref = sHref & "&Tabla=Eco&volver=" & vuelta & "&Eliminar=NSNC"
			Celda = ofv.BotonEliminar(rs(1),sHref)
		Else
			Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
		End If

		'id
		sInput = rs(0)
		Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")

		If session("APLMDLOR") = "BE"  Then
			SINPUT="<input name='"& rs(2).name   &"'  value='" & Trim(rs(2)&"") & "' type=text MAXLENGTH=2 class=dtn onblur="&Chr(34)&"validarnum(this,'"&sForm&"','"&Trim(rs(2)&"")&"',2);"&Chr(34) & ">"
			Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")
		End If




		'Descripcion
		If rs("estado") = 3 Then

		'	If session("UN") <> 1 or Session("EMPMUN") = "N" Then
				sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dtbs","")
		'	Else
		'		sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"80","dtbs","")
		'	End If

		Else

		'	If session("UN") <> 1 or Session("EMPMUN") = "N" Then
				sinput=rs(1)
		'	Else
		'		sinput=rs(11) & "&nbsp;-&nbsp;" & rs(1)
		'	End If

		End If

		'If session("UN") <> 1 or Session("EMPMUN") = "N" Then
			Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
		'Else
		'	Celda = Celda & ofv.gencelda("","vti","","20%","","","",sInput,"si")
		'End If


	'	If session("UN") = 1 And session("EMPMUN") = "S" Then
	'		Celda = Celda & ofv.gencelda("","vti","","30%","","","",rs(11),"si")
	'	End If

		'Propiedades


		If (session("APLMDLOR") & "") = "BE" Then

			sinput = "&nbsp;---"

		Else

			'sImagenP = "&#0232;"
			'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
			sinput = "<a href='SectoresProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
			sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
			sinput = sinput  & "' title='Propiedades de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
			If rs("estado") = 3 Then
				Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")
			Else
				Celda = Celda & ofv.gencelda("","ft","center","","15%","","","&nbsp;","si")
			End If
			SCOMSEP = Chr(9)
			scomK = "3" & SCOMSEP & "4"
			scomD = "VIGENTE" & SCOMSEP & "NO&nbsp;VIGENTE"
			sInput = ofv.gencomboFX(sForm,"ESTADO",RS("ESTADO"),scomK,scomD,"")

			'Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

		End If

		    If session("UN") <>  1 And session("EMPMUN") = "S" Then
		      sql1 ="select eco,descri from eco where estado=3 and ecoun =1 order by descri "
		      sql1 = "Select a.eco, replace('<b>' + b.cliente + ' : </b>' + a.descri,' ' ,'&nbsp;') From ECO a inner join clienteplan b on (a.ecoUN=b.ID)where a.ecoun=1 and a.estado=3  order by a.descri"
             '  response.write sql1
		       sInputS = ofv.GenerarCombo(sForm,rs(9).name,rs(9).value,sql1,cn,0,1,"")
               celda = Celda & ofv.gencelda("","ft","","","60%","","",sInputS,"si")
            end if


		    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

		    Fila = ofv.genrow("","","","","","","",Celda,"si")



		sAccion = ""
		If bModificar Then
			sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
			sAccion = sAccion & "&tabla=Eco&volver=" & vuelta
		End If
		Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

		posi = posi + 1
		rs.Movenext
	Loop
	
  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")


	retornar = ""
	If Session("EMPMUN") = "S" Then
		sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
		sInput = ofv.gencelda("","","center","","","","",sInput,"si")
		sAccion = "Sectores.asp"
		retornar = ofv.genform("Volver","",sAccion,"","post","",sInput,"si")

		sAccion = "Sectores.asp"
		retornar = ofv.BotonVolver(sAccion)
	End If


	'Botonera
	Agregar = ""
	If bAgregar Then
		sAccion = "AgregarSector.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=Eco"
		Agregar = ofv.BotonAgregar(sAccion)
	End If

	If session("UN") <> 1 or Session("EMPMUN") = "N" Then
		str1 = "Select descri From ECO where ecoun = " & SUN & " order by descri"    ' orden pagina
		str2 = "Select descri From ECO where ecoun = " & SUN & " order by descri"    ' orden alfabetico
	Else
		str1 = "Select a.descri + ' - ' + b.cliente From ECO a inner join clienteplan b on (a.ecoUN=b.ID) order by b.cliente"    ' orden pagina
		str2 = "Select a.descri + ' - ' + b.cliente From ECO a inner join clienteplan b on (a.ecoUN=b.ID) order by b.cliente"    ' orden alfabetico
	End If

	Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

	Listar  = ""

	If (session("APLMDLOR") & "") = "BE" Then

	Else
		'Listar Organigrama

		xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Listar el organigrama", "NO", "Organigrama")
		sAccion = "LisOrganigrama.asp"
		Listar  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")

	End If

	'consulta

	sAccion = "SectoresCons.asp?EMPRE=" & session("UN")
	''Celda = ofv.BotonConsultar(sAccion)
	Reportes=BtnReportes(16,"","xform",sAccion)

	sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Listar,Reportes & retornar)

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	shtml=SHTML & 	ValidarInput()
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
				sHref = "sectoresUN.asp?EMPRE=" & ATREMP(IAEMP,0)
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

Response.Write sHTML

%>

