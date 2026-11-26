<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  Tabla       = request.querystring("Tabla")
  vuelta      = request.querystring("vuelta")
  Regisvuelta = request.querystring("Regisvuelta")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR IMPACTO DE RIESGO")

  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","VALOR","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","10","dtN","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(2).name,"","agregar","","70","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")





  sFuncion2 = "onsubmit=" & chr(34) & "alert('Los campos OK');" & chr(34)

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0 || this." & rs(2).name & ".value.length==0 ) "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {alert('Los campos OK');}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../ORNIVELIMPACTO.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.genform("Agregarb","",sAccion,"","post",sFuncion2,Fila,"no") & chr(13)

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si") & chr(13)

  'Aceptar y Cancelar
  sAccion = "ORNIVELIMPACTO.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self") & chr(13)

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