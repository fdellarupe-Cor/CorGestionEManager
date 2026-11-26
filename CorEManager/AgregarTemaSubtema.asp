<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	
	Tabla = request.querystring("Tabla")
	vuelta = request.querystring("vuelta")
	Regisvuelta = request.querystring("Regisvuelta")
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	strsql = "Select * From " & tabla
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	
	
	If Session("CorApli") = "COREM20" Then
		wkaplic = "EM"
	ElseIf Session("CorApli") = "COREP20" Then
		wkaplic = "PL"
	Else
		wkaplic = "OR"
	End If
	
	
	sHTML = ofv.FormHeader("AGREGAR TEMA")
	
	sForm = "Agregarb"
	
	sInput = "AGREGAR " & session("formn")
	Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
	Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	
	Celda  = ofv.gencelda("","tt","","20%","","","","TEMA","si")
	sInput = ofv.GenerarInput(rs(1).name,"NUEVO TEMA","agregar","","140","dt","")
     IF SESSION("BASE")="K" THEN

     sInput = sinput & ofv.GenerarInput(rs(3).name,"K","hidden","","1","dt","")
     ELSE
	sInput = sinput & ofv.GenerarInput(rs(3).name,"C","hidden","","1","dt","")
     END IF
	sInput = sinput & ofv.GenerarInput(rs(5).name,wkaplic,"hidden","","1","dt","")
	Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	
	
  '  sFuncion = "onsubmit=" & chr(34) & "if (this.descri"
  '  sFuncion = sFuncion & ".value.length==0 ) "
  '  sFuncion = sFuncion & "{alert('Descripcion no debe ser nulo');return false;} " & chr(34)
       sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(1).name
      sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.trim()=='' "
      sFuncion = sFuncion & "|| fr." & rs(1).name & ".value==' ') "
    sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)

	sAccion = "ASP/AgregarTema.asp?Volver=../TEMASSUBTEMAS.asp?CantRegMover=" & Regisvuelta & "&Agregar=True"
    Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post","",Fila,"no")

	
    sHTML = sHTML & ofv.GenTabla("","aaa ","","80%","","0","","0","1","","",Formulario,"si")
	
	'Aceptar y Cancelar
	sAccion = "TEMASSUBTEMAS.asp?CantRegMover=" & Regisvuelta
    'sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")
  zboton=ofv.AgregarBoton("Agregarb",sAccion,"") & chr(13)
     zboton=replace(zboton,"'document.Agregarb.submit();'",sFuncion)
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