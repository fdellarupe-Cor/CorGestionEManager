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

  sHTML = ofv.FormHeader("AGREGAR cantidad copias")


  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","Tiempo de Guarda","si")
     sInput = ofv.GenerarInput(rs(0).name,"","agregar","","20","dtn","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
   '' Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Cantidad","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt","")
      sInput2 = ofv.GenerarInput("ESTADO","3","HIDDEN","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput & sInput2,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
  sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0) "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} "
  sFuncion = sFuncion & "else {if (isNaN(this." & rs(0).name & ".value)) "
  sFuncion = sFuncion & "{alert('Nivel debe ser numerico'); this." & rs(0).name & ".focus(); "
  sFuncion = sFuncion & "this." & rs(0).name & ".select();return false;}; };" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../FormAdminParcopias.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")


  'Aceptar y Cancelar
  sAccion = "FormAdminParcopias.asp?CantRegMover=" & Regisvuelta
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