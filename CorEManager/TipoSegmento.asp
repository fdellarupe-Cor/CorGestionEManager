<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	
	CantRegMostrarx = 8
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	
	sHTML = ofv.FormHeader(session("FormN"))
	
	'Encabezados
	If session("APLMDCLI")="BPN" Then
		Formulario = ofv.Encabezados("Id","Descripcion","secuencia","estado")
	Else
		Formulario = ofv.Encabezados("Id","Descripcion","estado")
	End If
	
	
	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	Strsql = "Select * From TipoSegmento Order By SECUEN"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then rs.move(CantRegMoverx)
	posi = 0
	Do While (Not rs.EOF) And (posi < CantRegMostrarx)
		sForm = "form" & posi
		
		'Eliminar
		If bSuprimir Then
			sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
			sHref = sHref & "&Tabla=TipoSegmento&volver=" & vuelta & "&Eliminar=NSNC"
			Celda = ofv.BotonEliminar(rs(1),sHref)
			
		Else
			Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
		End If
		
		'id
		sInput = rs(0)
		Celda = Celda & ofv.gencelda("","cc","right","50","","","",sInput,"si")
		
		
		'Descripcion
		sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
		Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
		
		'secuencia
		If session("APLMDCLI")="BPN" Then
			sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"25","dt","")
			Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
		End If
		SCOMSEP = Chr(9)
		scomK = "3" & SCOMSEP & "4"
		scomD = "VIGENTE" & SCOMSEP & "NO&nbsp;VIGENTE"
		If bModificar Then
			sInput = ofv.gencomboFX(sForm,"ESTADO",RS("ESTADO"),scomK,scomD,"")
		Else
			If  RS("ESTADO")= 3 Then
				sInput="VIGENTE"
			Else
				sInput="NO VIGENTE"
			End If
		End If
		Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
		Fila = ofv.genrow("","","","","","","",Celda,"si")
		
		sAccion = ""
		If bModificar Then
			sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
			sAccion = sAccion & "&tabla=TipoSegmento&volver=" & vuelta
		End If
		Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
		
		posi = posi + 1
		rs.Movenext
	Loop

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")
	
	'Botonera
	Agregar = ""
	If bAgregar Then
		sAccion = "AgregarTipoSegmento.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=TipoSegmento"
		Agregar = ofv.BotonAgregar(sAccion)
		
	End If
	
	
	If session("APLMDCLI")="BPN" Then
		str1 = "Select descri From tipoSegmento Order By SECUEN "    ' orden pagina
		str2 = "Select descri From tipoSegmento Order By SECUEN "    ' orden alfabetico
	Else
		str1 = "Select descri From tipoSegmento Order By DESCRI "    ' orden pagina
		str2 = "Select descri From tipoSegmento Order By DESCRI "    ' orden alfabetico
		
	End If
	
	Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
	
	
	'consulta
	
	
	
	sinput = "<a  href=TipoSegmentoCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
	Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
	
	sAccion = "TipoSegmentoCons.asp?" & request.querystring
	''Celda = ofv.BotonConsultar(sAccion)
	Reportes=BtnReportes(16,"","xform",sAccion)
	
	''sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
	sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,"no",Reportes)
	
	
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	
	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>