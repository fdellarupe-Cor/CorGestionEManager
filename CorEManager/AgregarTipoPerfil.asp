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

  sHTML = ofv.FormHeader("AGREGAR TIPO DE PERFIL")


  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'ndsperfil = clng(date()) & replace(time(),":","")
   ''response.write ndsperfil & " ndsperfil " & chr(13)
   ndsperfil = ndsperfil & clng(date()) &  time()
    ndsperfil=retnum (ndsperfil)


     Celda  = ofv.gencelda("","tt","","20%","","","","Perfil","si")
    'sInput = ofv.GenerarInput(rs(0).name,ndsperfil,"agregar","","20","dt","")
    'Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    sinput3 =ofv.GenerarInput(rs(0).name,ndsperfil,"HIDDEN","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",ndsperfil,"si")
     Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt","")
      sInput2 = ofv.GenerarInput("ESTADO","3","HIDDEN","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput & sInput2 & sInput3,"si")
     Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")




  'sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
  'sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0 "
  'sFuncion = sFuncion & "|| this." & rs(0).name & ".value==' ') "
  'sFuncion = sFuncion & "{alert('Los campos no deben ser nulos'); return false;}" & chr(34)

  'sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(0).name
  'sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.length==0 "
  'sFuncion = sFuncion & "|| fr." & rs(0).name & ".value==' ') "
  'sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)

sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(1).name
   sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value=='' "
   sFuncion = sFuncion & "|| fr." & rs(1).name & ".value==' ') "
   sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)

  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../TiposdePerfil.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post","",Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")



  'Aceptar y Cancelar
  sAccion = "TiposdePerfil.asp?CantRegMover=" & Regisvuelta
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

Response.Write sHTML

%>