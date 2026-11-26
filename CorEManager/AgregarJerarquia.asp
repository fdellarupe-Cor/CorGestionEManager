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

  sHTML = ofv.FormHeader("AGREGAR JERARQUIAS")

  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'ndsjerar = rs(0).name
    ndsjerar = ndsjerar & clng(date()) &  time()
    ndsjerar=retnum (ndsjerar)



    'ndsjerar = ndsjerar & "agregar"
  ''response.write ndsjerar & " ndsjerar " & chr(13)

     Celda  = ofv.gencelda("","tt","","20%","","","","JERARQUIA","si")
    'sInput = ofv.GenerarInput(rs(0).name,ndsjerar,"","","50","dt","")

   '' Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
        sinput3 =ofv.GenerarInput(rs(0).name,ndsjerar,"HIDDEN","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",ndsjerar,"si")
     Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt","")
     sInput2 = ofv.GenerarInput("ESTADO","3","HIDDEN","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput & sInput2 & sinput3 ,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

 'sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name & ".value.length==0 "
 'sFuncion = sFuncion & "|| this." & rs(1).name & ".value.length==0 ) "
 'sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} " & chr(34)


 sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(1).name
  sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.trim()=='' "
  sFuncion = sFuncion & "|| fr." & rs(1).name & ".value==' ') "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)

  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../Jerarquias.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post","",Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario,"si")


  'Aceptar y Cancelar
  sAccion = "Jerarquias.asp?CantRegMover=" & Regisvuelta
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