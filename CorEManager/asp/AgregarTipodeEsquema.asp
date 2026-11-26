<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	
	Dim nivel(10)
	nivel(0) = "C"
	nivel(1) = "D"
	nivel(2) = "E"
	nivel(3) = "F"
	nivel(4) = "G"
	nivel(5) = "H"
	nivel(6) = "I"
	nivel(7) = "J"
	nivel(8) = "K"
	nivel(9) = "L"
	
	
	
	
	Tabla = request.querystring("Tabla")
	vuelta = request.querystring("vuelta")
	Regisvuelta = request.querystring("Regisvuelta")
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	strsql = "Select * From " & tabla
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	
	sHTML = ofv.FormHeader("AGREGAR TIPOS DE ESQUEMAS NORMATIVOS")
	
	sForm = "Agregarb"
	
	sInput = "AGREGAR " & session("formn")
	Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
	
	Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	Celda  = ofv.gencelda("","tt","","20%","","","","Tipo","si")
	sInput = ofv.GenerarInput(rs(0).name,"","agregar","","20","dt","")
	Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
	sInput = ofv.GenerarInput(rs(1).name,"","agregar","","80","dt","")
	Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	Celda  = ofv.gencelda("","tt","","","","","","Perfil","si")
	Sql1   = "Select Tipoperfil,DESCRI From Objperfil Order By Tipoperfil"
	''' sInput = ofv.GenerarCombox3(Sql1,cn,0,0,rs(3).name)
	sInput = ofv.GenerarCombo(sForm,rs(3).name,"OBJMAIN",Sql1,cn,0,1,"no")
	Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	'Celda  = ofv.gencelda("","tt","","","","","","Icono","si")
	'sImagenV = ucase("ICorAbierto.jpg")
	'sImagenC = rs(15).name
	'sInput = ofv.GenerarCombo333(ofv.Getfiles("Iconos"),sImagenC,sImagenV,"Iconos\","imagen")
	''  sInput = ofv.GenComboM("9", sForm, sImagenC, sImagenV, "", "", "", "", ofv.Getfiles("Iconos"), "", "", "", "", "", "", "")
	'Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
	'Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	Celda  = ofv.gencelda("","tt","","","","","","Nivel","si")
	SCOMSEP = ""
	scomK   = ""
	
	For i=0 To 9
		scomK = scomK & SCOMSEP & nivel(i)
		SCOMSEP = Chr(9)
	Next
	
	scomD = scomK
	
	sInput = ofv.gencomboFX(sForm,rs(4).name,"C",scomK,scomD,"no")
	Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	
	
	
	
	
	sInput   = ofv.GenerarInput(rs(18).name,"8","HIDDEN","","1","dt","")
	sInput   = sInput &  ofv.GenerarInput(rs(16).name,"S","HIDDEN","","1","dt","")
	sInput   = sInput &  ofv.GenerarInput(rs(17).name,"S","HIDDEN","","1","dt","")
	sInput   = sInput &  ofv.GenerarInput(rs(21).name,"D","HIDDEN","","1","dt","")
	sInput   = sInput &  ofv.GenerarInput(rs(22).name,"S","HIDDEN","","1","dt","")
	
	If SESSION("SETPROC") = "ALL" Then
		sInput   = sinput &  ofv.GenerarInput(rs(23).name,wkaplic,"hidden","","1","dt","")
	Else
		sInput   = sinput &  ofv.GenerarInput(rs(23).name,SESSION("SETPROC"),"hidden","","1","dt","")
	End If
	
	For i=0 To 9
		sInput  = sInput &  ofv.GenerarInput(rs(5 + i).name,"1","HIDDEN","","1","dt","")
	Next
	
	sInput2 = ofv.GenerarInput(rs(19).name,"WKF00016631121722600921928","HIDDEN","","1","dt","")
	
	sInput3 = ofv.GenerarInput(rs(20).name,"NO","HIDDEN","","1","dt","")
	sInput4 = ofv.GenerarInput("avisoant","5","HIDDEN","","1","dt","")
	sInput5 = ofv.GenerarInput("textoimpresion","NO","HIDDEN","","1","dt","")
	'   sInput2 = ofv.GenerarInput(rs(19).name,"WKF00016631121722600921928999","HIDDEN","","1","dt","")
	Celda   = ofv.gencelda("","vt colspan=2","","","","","",sInput & sInput2 & sInput3 & sInput4 & sInput5,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	sFuncion2 = "onsubmit=" & Chr(34) & "alert('Los campos OK');" & Chr(34)
	
	sFuncion = "onsubmit=" & Chr(34) & "if (this." & rs(0).name
	sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0 || "
	sFuncion = sFuncion & "this." & rs(3).name & ".value.length==0 || "
	sFuncion = sFuncion & "this." & rs(0).name & ".value==' ') "
	sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {alert('Los campos OK');}" & Chr(34)
	sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../TiposdeEsquema.asp?CantRegMover="
	sAccion = sAccion & Regisvuelta
	Formulario = Formulario & ofv.genform("Agregarb","",sAccion,"","post",sFuncion2,Fila,"no") & Chr(13)
	
	sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si") & Chr(13)
	
	'Aceptar y Cancelar
	sAccion = "TiposdeEsquema.asp?CantRegMover=" & Regisvuelta
	sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self") & Chr(13)
	
	Agregar = ""
	Buscar  = ""
	Celda   = ""
	
	sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
	
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	
	sHTML = sHTML & "</body></html>"
End If

response.write sHTML

%>