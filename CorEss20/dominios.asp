
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
	
	sHTML = ofv.FormHeader(session("FORMN"))
	
	'Encabezados
	
	' if session("usrprv")<>"I" then
	
	If session("empmun") ="S" Then
		Formulario = ofv.Encabezados("Cod","descripcion","Raiz ldap","usuario","pwd","cliente")
		
	Else
		Formulario = ofv.Encabezados("Cod","descripcion","Raiz ldap","usuario","pwd")
	End If
	
	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar
	
	Set cn = ofv.conectar(ofv.strconn0)
	
	Strsql = "Select * From dominios Order By id "
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then rs.move(CantRegMoverx)
	posi = 0
	Do While (Not rs.EOF) And (posi < CantRegMostrarx)
		sForm = "form" & posi
		
		'Eliminar
		If bSuprimir  Then
			sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
			sHref = sHref & "&Tabla=dominios&volver=" & vuelta & "&Eliminar=NSNC"
			Celda = ofv.BotonEliminar(rs(0),sHref)
			rd=""
		Else
			rd="NO"
			Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
		End If
		
		'codigo
		sInput = rs(0)
		
		Celda = Celda & ofv.gencelda("","vt","","5%","","","",sInput,"si")
		
		
		'Descripcion
		sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"60","dt","")
		
		Celda = Celda & ofv.gencelda("","","","30%","","","",sInput,"si")
		'if session("usrprv")<>"I" then
		'RAIZLDAP
		
		If rD<> "NO" Then
			sInput = ofv.GenerarInput(rs(2).name, rs(2),"text",sForm,"60","dt","")
		Else
			sInput =RS(2)
		End If
		
		
		
		Celda = Celda & ofv.gencelda("","","","10%","","","",sInput,"si")
		
		If rD<> "NO" Then
			sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"50","dt","")
			
		Else
			sInput=RS(3)
		End If
		
		Celda = Celda & ofv.gencelda("","","","10%","","","",sInput,"si")
		' ccc=cstr(rs(4)&""))
		'response.write (rs(4)&"" ) <> "" = "" and trim(len(rs(4))) > 0
		
		If rD<> "NO" Then
			If (rs(4)&"" ) <> "" And Trim(Len(rs(4)&"")) > 0  And IsNumeric(rs(4)&"")Then
				ccc=CStr(ofv.getCript("D",rs(4)&""))
			Else
				ccc=rs(4)
			End If
			sInput = ofv.GenerarInput(rs(4).name,ccc,"password",sForm,"50","dt  onblur=""document." & sform & ".submit();""","")
		Else
			sInput=RS(4)
		End If
		
		
		
		Celda = Celda & ofv.gencelda("","","","10%","","","",sInput,"si")
		
		If session("empmun") ="S" Then
			strsql = "Select * From clienteplan Order By cliente"
			sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5),strsql,cn,0,1,"")
			'Descripcion
			'sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"80","dt","")
			Celda = Celda & ofv.gencelda("","","","50%","","","",sInput,"si")
			
		End If
		'   end if
		Fila = ofv.genrow("","","","","","","",Celda,"si")
		
		sAccion = ""
		If bModificar Then
			sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=dominios"
		End If
		Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
		
		posi = posi + 1
		rs.Movenext
	Loop
	ftab = ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")
	
	'Botonera
	Agregar = ""
	If bAgregar Then
		sAccion = "Agregardominios.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=dominios"
		Agregar = ofv.BotonAgregar(sAccion)
	End If
	
	str1 = "Select descripcion From dominios Order By descripcion"    ' orden pagina
	str2 = str1 '"Select observaciones From feriados Order By observaciones "    ' orden alfabetico
	Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
	
	sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")
	
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	
	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>