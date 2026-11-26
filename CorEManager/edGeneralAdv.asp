<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	bModificaresp = ofv.AModificar
	If session("Nomodificar") = "S" Then

		If Left(Session("Form2"),4) <> "DESA" Then    bModificar = False

	End If

	nCodigo = request.querystring("Codobjeto")
	sTipo   = request.querystring("Tipobjeto")
	sOPCION = request.querystring("OPCION")
	vuelta  = ofv.vueltaasp2("../")

	sHTML = ofv.FormHeader(session("FormN"))

	'----------------------------------------------

	shtml = shtml & CorPnlFnc("N","")

	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

	'-------------------------------------------------

	'Encabezados
	' Formulario = ofv.Encabezados("AGRUPAMIENTO","SEGMENTO")
	Set cn = ofv.conectar(ofv.strconn2)

	If session("APLMDCLI") = "BPN" Then

		sql = "select a.descri from tipoproducto a "
		sql = sql & " inner join objetos b on (a.objid = b.tsegmento) "
		sql = sql & " where b.codobjeto =" & nCodigo & " and b.tipobjeto='" & sTipo & "'"
		Set rx = ofv.crearconsultaEx(sql, cn, 1, parametros)

		prod = ""
		flag = False

		If Not rx.eof Then
			prod = rx(0)
			flag = True
		End If

		Call ofv.cerrarconsulta(rx)

		Formulariocabe = ""

		If flag Then
			If session("APLMDCLI") = "BPN" Then
				celdas = ofv.GenCelda("","ttr colspan =3","","2%","","","","Linea de negocio: " & prod,"si")
			Else
				celdas = ofv.GenCelda("","ttr colspan =3","","2%","","","","PRODUCTO: " & prod,"si")
			End If
		Else
			If session("APLMDCLI") = "BPN" Then
				celdas = ofv.GenCelda("","ttr colspan =3","","2%","","","","PARA AGREGAR PROCESO Y SUBPROCESO PRIMERO DEBE CARGAR LINEA DE NEOGOCIO","si")
			Else
				celdas = ofv.GenCelda("","ttr colspan =3","","2%","","","","PARA AGREGAR AGRUPAMIENTO Y SEGMENTO PRIMERO DEBE CARGAR UN PRODUCTO","si")
			End If
		End If
		Formulariocabe = Formulariocabe & ofv.GenRow("","","","","20","","",celdas,"si")
	Else
		flag=True
		celdas = ofv.GenCelda("","ttr colspan =3","","2%","","","","","si")
		Formulariocabe = Formulariocabe & ofv.GenRow("","","","","20","","",celdas,"si")
	End If

	celdas = ofv.GenCelda("","ttr","","2%","","","","&nbsp;","si")
	If session("APLMDCLI") = "BPN" Then
		celdas = celdas & ofv.GenCelda("","ttr","","49%","","","","PROCESO","si")
		celdas = celdas & ofv.GenCelda("","ttr","","49%","","","","SUBPROCESO","si")
	Else
		celdas = celdas & ofv.GenCelda("","ttr","","49%","","","","AGRUPAMIENTO","si")
		celdas = celdas & ofv.GenCelda("","ttr","","49%","","","","SEGMENTO","si")
	End If
	Formulariocabe = Formulariocabe & ofv.GenRow("","","","","20","","",celdas,"si")

	sImagenI   = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"

	' set cn = ofv.conectar(ofv.strconn2)

	If  SOPCION = "SI" Then
		Strsql = "insert into ObjTSEGAGR  (codobjeto,tipobjeto,SECUEN,TSEGMENTO,TAGRUPA) values ("
		Strsql = Strsql & nCodigo & ",'" & stipo & "','10','0','0')"
		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)
	End If

	Strsql = "Select * From ObjTSEGAGR Where codobjeto=" & nCodigo & " Order By OBJID"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	Do  Until rs.eof
		sForm = "edP" & rs(0)

		If  bModificar Then
			sAccion    = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=ObjTSEGAGR&volver=" & vuelta
			sAccion    = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
		End If

		'Eliminar
		If  bModificar Then
			sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=ObjTSEGAGR&volver=" & vuelta
			sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
		End If

		If  sHref <> "" Then
			Celda = ofv.BotonEliminar(rs(0),sHref)
		Else
			Celda = ofv.GenCelda("","","","","","","","","si")
		End If

		'AGRUPAMIENTO

		Strsql = "Select * From TIPOAGRUPA where estado=3 order by descri"
		''  sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,Strsql,cn,0,1,"")

		Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
		scomK = ""'"null" & Chr(9)

		scomD = "" '"" & Chr(9)
		SCOMSEP =""
		Do While Not rsperfil.eof
			scomK=scomK & SCOMSEP & rsperfil(0)
			scomd=scomd & SCOMSEP & rsperfil(1)
			SCOMSEP = Chr(9)
			rsPerfil.movenext
		Loop

		sInput =  ofv.gencomboFX(sForm & "w1", rs(5).name, rs(5).value, scomK, scomD, "")
		sInput = ofv.GenForm(sForm & "w1","",sAccion,"","post","",sInput,"si")
		'Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",sinput,"si")
		'Celda  = Celda & ofv.GenCelda("","VT","","40%","","","",formulario,"si")
		Celda  = Celda & ofv.GenCelda("","VT","","40%","","","",sInput,"si")

		'SEGMENTO

		Strsql = "Select * From TIPOSEGMENTO where estado=3 order by descri "
		'sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,Strsql,cn,0,1,"")
		Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
		scomK = ""'"null" & Chr(9)

		scomD = "" '"" & Chr(9)
		SCOMSEP =""
		Do While Not rsperfil.eof
			scomK=scomK & SCOMSEP & rsperfil(0)
			scomd=scomd & SCOMSEP & rsperfil(1)
			SCOMSEP = Chr(9)
			rsPerfil.movenext
		Loop

		sInput =  ofv.gencomboFX(sForm & "w2", rs(4).name, rs(4).value, scomK, scomD, "")
		sInput = ofv.GenForm(sForm & "w2","",sAccion,"","post","",sInput,"si")
		'  Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",sinput,"si")
		'Celda  = Celda & ofv.GenCelda("","VT","","40%","","","",formulario,"si")
		Celda  = Celda & ofv.GenCelda("","VT","","40%","","","",sInput,"si")


		'' Fila   = ofv.GenRow("","","","","","","",Celda,"si")

		' ' if  bModificar then
		' '     sAccion    = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=ObjTSEGAGR&volver=" & vuelta
		' '     sAccion    = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
		' ' ''    Formulario = ofv.GenTabla("","","","80%","","0","","-2","-1","","",fila,"si")
		' '     xFormulario =  ofv.GenForm(sForm,"",sAccion,"","post","",celda,"si")
		' '  ''   Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",formulario,"si")
		' ' else
		' '   ''  Formulario = ofv.GenTabla("","","","80%","","0","","-2","-1","","",fila,"si")
		' '    xFormulario =  ofv.GenForm("","","","","post","",celda,"si")
		' '    '' Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",formulario,"si")
		' ' end if
		Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")


		rs.movenext
	Loop

	sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulariocabe & Formulario,"si")
	''  tabla= ofv.GenTabla("","","","70%","","0","","0","0","","",Formulariocabe & fila,"si")
	''  sHTML = sHTML & tabla
	'' sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","-2","-1","","",Formulario,"si")
	'' sHTML = sHTML & formulario

	Agregar = ""

	'Botonera
	If  bModificar  And flag Then
		sAccion = "EDGENERALADV.asp?OPCION=SI&Codobjeto=" & nCodigo
		sAccion = sAccion & "&Tipobjeto=" & sTipo
		Agregar = ofv.BotonAgregar(sAccion)
	End If

	sAccion = "edGENERAL.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
	sRetornar = ofv.BotonVolver(sAccion)



	'--------------------------------------------------------------------------------
	sHTML = sHTML & "</span></div>"

	Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")

	'--------------------------------------------------------------------------------

	sHTML   = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,sRetornar,"")

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)

	sHTML = sHTML & " </body></html>"
End If

response.write sHTML

%>

