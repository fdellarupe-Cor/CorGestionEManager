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

  sHTML = ofv.FormHeader("AGREGAR PALABRA/FRASE AL GLOSARIO")


  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = "&nbsp;"
    Celda = ofv.gencelda("","ut COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","PALABRA/FRASE","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","100","dt  onblur='document.getElementById('" & rs(2).name & "').value = document.getElementById('" & rs(1).name & "').value;' ","")
     sInput = SINPUT & ofv.GenerarInput(rs(2).name,"NO","hidden","","100","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0  "
  sFuncion = sFuncion & "|| this." & rs(1).name & ".value==' ') "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} ELSE "
  sFuncion = sFuncion & "{this." & rs(2).name & ".value = this." & rs(1).name & ".value;}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../GLOSMANAGER.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.genform("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")

  'Aceptar y Cancelar
  sAccion = "GLOSMANAGER.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>