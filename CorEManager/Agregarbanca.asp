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

  sHTML = ofv.FormHeader("AGREGAR BANCA")


  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")




     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(2).name,"","agregar","","120","dt","")
      sInput2 = ofv.GenerarInput("ESTADO","3","HIDDEN","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput & sInput2,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

'  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(2).name
'  sFuncion = sFuncion & ".value.length==0 || this." & rs(2).name & ".value.length==0) "
'  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} "
'  sFuncion = sFuncion & "else {if (isNaN(this." & rs(2).name & ".value)) "
'  sFuncion = sFuncion & "{alert('Nivel debe ser numerico'); this." & rs(2).name & ".focus(); "
'  sFuncion = sFuncion & "this." & rs(2).name & ".select();return false;}; };" & chr(34)
 ' sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../FormAdminParBANCA.asp?CantRegMover="
 ' sAccion = sAccion & Regisvuelta

  sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(2).name
    sFuncion = sFuncion & ".value.length==0 || fr." & rs(2).name & ".value.length==0 "
    sFuncion = sFuncion & "|| fr." & rs(2).name & ".value==' ') "
    sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)
    sAccion = "asp/Agregar.asp?Tabla=" & Tabla
    sAccion = sAccion & "&volver=../FormAdminParBANCA.asp?CantRegMover=" & Regisvuelta
  Formulario = Formulario & ofv.genform("Agregarb","",sAccion,"","post","",Fila,"no")


  'Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no") & chr(13)

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")


  'Aceptar y Cancelar
  sAccion = "FormAdminParBANCA.asp?CantRegMover=" & Regisvuelta

  '	 sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")
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