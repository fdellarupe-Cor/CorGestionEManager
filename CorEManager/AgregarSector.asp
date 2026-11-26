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

  sHTML = ofv.FormHeader("AGREGAR SECTORES")

  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","DESCRIPCION","si")
     sInput = ofv.GenerarInput(rs(1).name,"NUEVO SECTOR","agregar","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ""
  'Codigo interno
  sInput1 = ofv.GenerarInput(rs(2).name,"0","hidden","","1","dt","")
  'Segmento
  sInput2 = ofv.GenerarInput(rs(4).name,"0","hidden","","1","dt","")
  'Nivel
  sInput3 = ofv.GenerarInput(rs(5).name,"0","hidden","","1","dt","")
  'Jerarquia
  sInput4 = ofv.GenerarInput(rs(6).name,"*","hidden","","1","dt","")
  'UN
  sInput5 = ofv.GenerarInput(rs(7).name,session("UN"),"hidden","","1","dt","")
  'Site
  sInput6 = ofv.GenerarInput(rs(8).name,"1","hidden","","1","dt","")

  sInput7 = ofv.GenerarInput("estado","3","hidden","","1","dt","")

  Celda = Celda & ofv.gencelda("","ut","","","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4 & sInput5 & sInput6 & sInput7,"si")


  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  'sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".value.length==0) "
  'sFuncion = sFuncion & "{alert('El campo Descripcion no debe ser nulo');"
  'sFuncion = sFuncion & "return false;}"
  'sFuncion = sFuncion & " else {if (this." & rs(2).name & ".value.length==0)"
  'sFuncion = sFuncion & " {this." & rs(2).name & ".value = 0;} return true;}" & chr(34)
  sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(1).name
  sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.trim()=='' "
   sFuncion = sFuncion & "|| fr." & rs(1).name & ".value==' ') "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../SectoresUN.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta & "&Campo=1"
 ' Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

 ' sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

 Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post","",Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario,"si")

  'Aceptar y Cancelar
  sAccion = "SectoresUN.asp?CantRegMover=" & Regisvuelta & "&tabla=" & tabla
  'sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")
  zboton=ofv.AgregarBoton("Agregarb",sAccion,"") & chr(13)
    zboton=replace(zboton,"'document.Agregarb.submit();'",sFuncion)
  sHTML = sHTML & zboton

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