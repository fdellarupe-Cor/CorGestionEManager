<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  Regisvuelta = request.querystring("Regisvuelta")
  NOM = request.querystring("NOM")
  NOM = ofv.convertircar(NOM,"*"," ")
  ID = request.querystring("ID")
  Tabla = request.querystring("Tabla")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)

  sForm = "Agregarb"


    sInput = "AGREGAR Perfil en Sector: " & NOM
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","Perfil","si")
     Strsql = "Select * From TIPOPERFIL where estado=3 order by descri"
     sInput = ofv.GenerarCombo(sForm,rs(2).name,"",Strsql,cn,0,1,"no")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


     sInput = ofv.GenerarInput(rs(3).name,"1","hidden","","","","")
     sInput = sInput & ofv.GenerarInput(rs(1).name,ID,"hidden","","","","")
     Celda  = ofv.gencelda("","vt colspan=2","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")



  f1f = "onsubmit=" & chr(34) & "if (this." & rs(2).name
  f1f = f1f & ".value.length==0 || this." & rs(3).name & ".value.length==0) "
  f1f = f1f & "{alert('Los campos no deben ser nulos');return false;} "
  f1f = f1f & "else {if (isNaN(this." & rs(3).name & ".value)) "
  f1f = f1f & "{alert('Dotacion debe ser numerico'); this." & rs(3).name & ".focus(); "
  f1f = f1f & "this." & rs(3).name & ".select();return false;}; };" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=../SectorPerfiles.asp?CantRegMover=" & Regisvuelta & "&ID=" & ID
  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",f1f,Fila,"no")

 sHTML = sHTML & ofv.GenTabla("","aaa ","","80%","","0","","0","1","","",Formulario,"si")


  'Botones de aceptar/cancelar
  sAccion = "SectorPerfiles.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

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