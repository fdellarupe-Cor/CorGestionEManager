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

  sHTML = ofv.FormHeader("AGREGAR Perfiles de Tablas")


  sForm = "Agregarb"
  sFormX = "AgregarX"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","Perfil","si")
     sInput = ofv.GenerarInput(rs(0).name,"","text",sFormX,"40","dt","")
     sInput = REPLACE(sInput,"{ document." & sFormX  & ".submit(); return true;}","{ if (this.value.length > 10) {alert('El texto es demasiado extenso'); return false;} else { document." & sFormX & ".submit(); return true;}}")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sInput   = ofv.GenerarInput(rs(2).name,"1","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(7).name,"1","HIDDEN","","1","dt","")

  sInput   = sInput &  ofv.GenerarInput(rs(3).name,"#FFFFFF","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(4).name,"#000080","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(5).name,"TAHOMA","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(6).name,"9PT","HIDDEN","","1","dt","")

  sInput   = sInput &  ofv.GenerarInput(rs(8).name,"#993333","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(9).name,"#FFFFFF","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(10).name,"TAHOMA","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(11).name,"9PT","HIDDEN","","1","dt","")

  sInput   = sInput &  ofv.GenerarInput(rs(12).name,"#000080","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(13).name,"#FFFFFF","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(14).name,"TAHOMA","HIDDEN","","1","dt","")
  sInput   = sInput &  ofv.GenerarInput(rs(15).name,"8PT","HIDDEN","","1","dt","")

  sInput   = sInput &  ofv.GenerarInput(rs(16).name,"100","HIDDEN","","1","dt","")
  For i=0 to 9
      sInput  = sInput &  ofv.GenerarInput(rs(17 + i).name,"0","HIDDEN","","1","dt","")
  next
  Celda   = ofv.gencelda("","vt colspan=2","","","","","",sInput,"si")
  Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")



  'sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
  'sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0 "
  'sFuncion = sFuncion & "|| this." & rs(0).name & ".value==' ') "
  'sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} " & chr(34)

  sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(0).name
  sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.length==0 "
  sFuncion = sFuncion & "|| fr." & rs(0).name & ".value==' ') "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../perfilesdetablas.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post","",Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "perfilesdetablas.asp?CantRegMover=" & Regisvuelta
  'sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

   zboton=ofv.AgregarBoton("Agregarb",sAccion,"") & chr(13)
  zboton=replace(zboton,"'document.Agregarb.submit();'",sFuncion)
  sHTML = sHTML & zboton

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>