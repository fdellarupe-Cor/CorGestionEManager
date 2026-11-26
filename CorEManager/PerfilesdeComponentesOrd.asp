<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	
	CantRegMostrarx = request.querystring("CantRegMostrar")
	CantRegMoverx = request.querystring("CantRegMover")
	
	
	wperfil = request.querystring("wperfil")
	welem = request.querystring("welem")
	
	If Len(welem) = 0 Then welem = 0
	
	
	sHTML = ofv.FormHeader(session("FormN"))
	
	'Encabezados
	Formulario = ""
	
	
	Dim stabla(15)
	Dim stablaE(15)
	Dim stablaS(15)
	
	
	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	posi = 0
	
	
	Strsql = "Select * From comperfil where tipoperfil = '" & wperfil & "' "
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		
		fila = ""
		
		
		Celda = ofv.GenCelda("","tt  ","","","","","","Perfil: ","si")
		sInput = rs(0) & "&nbsp;-&nbsp;" & rs(1)
		
		Celda = celda & ofv.GenCelda("","cc  ","","90%","","","",sInput,"si")
		
		sinput = "<a  href=PerfilesdecomponentesProp.asp?CantRegMover=" & CantRegMoverx
		sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil & " >Volver</a>"
		Celda = Celda & ofv.gencelda("","bietx style='display:flex;margin-top: 0.4rem;' ","center","","15%","","",sInput,"si")
		Fila = ofv.genrow("","","","","","","",Celda,"si")
		
		
		Celda = ofv.GenCelda("","aaa colspan=3  ","","","","","","&nbsp;","si")
		
		Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
		
		
		sInput = "&nbsp;Orden&nbsp;de&nbsp;Solapas"
		
		Celda = ofv.GenCelda("","tt colspan=3 ","","","","","",sInput,"si")
		
		Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
		
		
		Formulario = Fila
		
		
		
		'      sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")
		
		'      Formulario = ""
		
		
		
		
		For xx = 2 To 16
			stabla (xx - 1) = rs(xx + 30)
		Next
		
		
		yy = 0
		For xx = 1 To 15
			If CLng(stabla (xx)) < 99 Then
				yy = yy + 1
				stablaE (stabla (xx)) = xx
			End If
		Next
		tyy = yy
		' aca ,modifique clng(welem) = 15
		If CLng(welem) = 16 Then welem = 0
		' response.write welem
		If CLng(welem) > 0 Then
			'       if  clng(stabla(clng(welem) - 1)) < 3 then welem = 0
		End If
		
		If CLng(welem) > 0 Then
			
			
			welemix = stabla(CLng(welem) - 1)
			welemixa = CLng(welemix) - 1
			welemixh = CLng(welemix) - 2
			
			For xx = 1 To welemixh
				
				stablaS(xx) =  stablaE(xx)
				
			Next
			
			stablaS(welemix) =  stablaE(welemixa)
			stablaS(welemixa) =  welem - 1
			
			For xx = (CLng(welemix) + 1) To tyy
				
				stablaS(xx) =  stablaE(xx)
				
			Next
			
			Strsql = "UPDATE comPERFIL SET "
			
			For xx = 1 To tyy
				
				Strsql = strsql & rs(stablaS(xx) + 1 + 30).name & " = " & xx
				If xx <> tyy Then Strsql = strsql & ", "
				
			Next
			
			Strsql = strsql & " where tipoperfil = '" & wperfil & "' "
			'response.write strsql
			Call ofv.crearconsultaEx(StrSql,cn,1,parametros)
			
			
		Else
			
			For xx = 1 To tyy
				
				stablaS(xx) =  stablaE(xx)
				
			Next
			
		End If
		
		
		xx = 0
		
		For yy = 1 To tyy
			
			
			fila = ""
			
			
			
			xx = xx + 1
			Celda = ofv.GenCelda("","cc  ","","","","","",xx & "º ","si")
			sInput = rs(stablaS(yy) + 1 + 15)
			
			If rs(stablaS(yy) + 1) Then
				
				
				If xx = 1 Or xx = 82 Then
					
					Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")
					
					
					Celda = Celda & ofv.gencelda("","tt style='background-color:#008000;' ","center","","15%","","","-----","si")
					
				Else
					
					Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")
					
					sinput = "<a  href=PerfilesdecomponentesOrd.asp?CantRegMover=" & CantRegMoverx
					sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil    & "&Welem=" & (stablaS(yy) + 1) & " >Subir</a>"
					Celda = Celda & ofv.gencelda("","bietx style='display:block;margin-top: 0.4rem;' ","center","","15%","","",sInput,"si")
					
				End If
				
			Else
				
				Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")
				
				
				Celda = Celda & ofv.gencelda("","tt style='background-color:#990000;' ","center","","15%","","","-----","si")
				
			End If
			
			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
			
			Formulario = Formulario & Fila
			
			
			
		Next
		
	End If
	sHTML = sHTML & ofv.GenTabla("","aaa","","80% ","","0","","0","2","","",Formulario,"si")
	
	
	
	
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	
	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>