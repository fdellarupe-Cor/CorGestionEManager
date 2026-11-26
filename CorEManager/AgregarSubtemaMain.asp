<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	
	Tabla       = request.querystring("Tabla")
	vuelta      = request.querystring("vuelta")
	Regisvuelta = request.querystring("Regisvuelta")
	sTema       = request.querystring("Tema")
	
	Set cn = ofv.conectar(ofv.strconn2)
	'response.write request.querystring
	strsql = "Select * From " & tabla
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	
	sHTML = ofv.FormHeader("AGREGAR JERARQUIAS")
	
	sForm = "Agregarb"
	
	sInput = "AGREGAR " & session("formn")
	Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
	Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	
	Celda  = ofv.gencelda("","tt","","20%","","","","SUBTEMA","si")
	sInput = ofv.GenerarInput(rs(1).name,"NUEVO SUBTEMA","agregar","","120","dt","")
	Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	Celda  = ofv.gencelda("","tt","","","","","","TIPO","si")
	If SESSION ("BASE") ="K" Then
		strsql = "Select tipobjeto,descri From Subtema where icono='K' OR TIPOBJETO='PEDIDO' Order By descri"
		
	Else
		strsql = "Select tipobjeto,descri From Subtema where icono <> 'K' Order By descri"
	End If
	Set rse = ofv.crearconsultaEx(StrSql,cn,1,volver)
	xdato=""
	If Not rse.eof Then
		xdato=rse(0)
	End If
	
	'RESPONSE.WRITE strsql
	sInput = ofv.GenerarCombo(sForm,rs(4).name,xdato,Strsql,cn,0,1,"no")
	Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	
	
	sInput = ofv.GenerarInput(rs(5).name,"0","hidden","","1","dt","")
	sInput = sInput & ofv.GenerarInput(rs(6).name,"0","hidden","","1","dt","")
	If SESSION ("BASE") ="K" Then
		
		sInput = sinput & ofv.GenerarInput(rs(7).name,"K","hidden","","1","dt","")
	Else
		sInput = sinput & ofv.GenerarInput(rs(7).name,"C","hidden","","1","dt","")
	End If
	Celda  = ofv.gencelda("","vti COLSPAN=2 ","","","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	
	'sFuncion = "onsubmit=" & Chr(34) & "if (this.descri"
	'sFuncion = sFuncion & ".value.length==0 ) "
	'sFuncion = sFuncion & "{alert('Descripcion no debe ser nulo');return false;} " & Chr(34)
	
	'' sAccion = "AgregarSubtema.asp?Volver=" & vuelta & "&Agregar=True"
	sFuncion =  Chr(34) & " var fr=document.Agregarb; if (fr." & rs(4).name
	sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.length==0 "
	sFuncion = sFuncion & "|| fr." & rs(4).name & ".value==' ') "
	sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & Chr(34)
	sAccion = "ASP/AgregarSubtema.asp?Tema=" & STema & "&Agregar=True&Volver=../SUBTEMAS.asp?CantRegMover=" & Regisvuelta
	Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")
	
	sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
	
	'Aceptar y Cancelar
	sAccion = "SUBTEMAS.asp?CantRegMover=" & Regisvuelta
	'sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")
	zboton=ofv.AgregarBoton("Agregarb",sAccion,"") & Chr(13)
	zboton=Replace(zboton,"'document.Agregarb.submit();'",sFuncion)
	sHTML = sHTML & zboton
	
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