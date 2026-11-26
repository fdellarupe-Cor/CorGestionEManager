<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Tabla = request.querystring("Tabla")
''response.write Tabla & " Tabla " & chr(13)
  vuelta = request.querystring("vuelta")
  Regisvuelta = request.querystring("Regisvuelta")

  set cn2 = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & Tabla
''response.write strsql & " strsql " & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn2,1,volver)

  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)

  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","DESCRIPCION","si")
     sInput = ofv.GenerarInput(rs(1).name,"NUEVO PRODUCTO","agregar","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda  = ofv.gencelda("","tt","","20%","","","","SECUENCIA","si")
	sInput = ofv.GenerarInput(rs(2).name,"0","agregar","","20","dt","")
    Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
  sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0) "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} "
  sFuncion = sFuncion & "else {if (isNaN(this." & rs(0).name & ".value)) "
  sFuncion = sFuncion & "{alert('Nivel debe ser numerico'); this." & rs(0).name & ".focus(); "
  sFuncion = sFuncion & "this." & rs(0).name & ".select();return false;}; };" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../TipoPRODUCTO.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

 sHTML = sHTML & ofv.GenTabla("","aaa ","","80%","","0","","0","1","","",Formulario,"si")


  'Aceptar y Cancelar
  sAccion = "TipoPRODUCTO.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>