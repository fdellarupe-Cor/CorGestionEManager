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
  vuelta = request.querystring("vuelta")
  Regisvuelta = request.querystring("Regisvuelta")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR TIPO DE NIVEL PUESTO")

  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","Nivel","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","20","dtn","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(2).name,"","agregar","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(2).name & ".value.length==0 || this."
  sFuncion = sFuncion & rs(1).name & ".value.length==0) "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} "
  sFuncion = sFuncion & "else {if (isNaN(this." & rs(2).name & ".value)|| this." & rs(2).name
  sFuncion = sFuncion & ".value < 0) {alert('Nivel debe ser numerico'); this." & rs(2).name
  sFuncion = sFuncion & ".focus(); this." & rs(2).name & ".select();return false;}; };" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../Nivelpuesto.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario,"si")



  'Aceptar y Cancelar
  sAccion = "NivelpuESTO.asp?CantRegMover=" & Regisvuelta
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