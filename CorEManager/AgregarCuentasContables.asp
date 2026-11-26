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
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR CUENTAS CONTABLES")

  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","CUENTA","si")
     sInput = ofv.GenerarInput(rs(2).name,"","agregar","","30","dt","")
     sInput2 = ofv.GenerarInput("ESTADO","3","HIDDEN","","10","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput & sInput2,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","DESCRIPCION","si")
     sInput = ofv.GenerarInput(rs(1).name,"NUEVA CUENTA","agregar","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0 || this." & rs(2).name & ".value.length==0) "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../CuentasContables.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta & "&Campo=2"
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario,"si")

  'Aceptar y Cancelar
  sAccion = "CuentasContables.asp?CantRegMover=" & Regisvuelta
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