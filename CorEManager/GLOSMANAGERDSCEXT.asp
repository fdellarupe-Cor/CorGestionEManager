<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  GLOSDSC = request.querystring("GLOSDSC")
  IF GLOSDSC <> "" THEN
     SESSION("GLOSDSC") = GLOSDSC
  ELSE
     GLOSDSC = SESSION("GLOSDSC")
  END IF 




  CantRegMostrarx = 0
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)



  set cn = ofv.conectar(ofv.strconn2)


  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)


  'Encabezados
  Celda      = ofv.GenCelda("","ttt colspan=2 ","","","","","","Descripcion","si")
  Formulario = ofv.genrow("","","","","22","","",Celda,"si")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  Strsql = "Select * From GLOSDSC WHERE OBJID = " & GLOSDSC
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 
  posi = 0

    sForm = "form" & posi


      Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")


      'Texto

    sInput = "<font class=at><textarea class=dt name=" & rs(4).name & "  id=" & rs(4).name & " rows=20 cols=90 "
    if not bModificar then
       sInput = sInput & " readonly "
    else
       sInput = sInput & " onkeypress=" & chr(34)
       sInput = sInput & " if (this.value.length > 1950) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;}" & chr(34)
       sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit() ;" & chr(34)
    end if
    sInput = sInput & " >" & rs(4) & "</textarea></font>"

    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=GLOSDSC&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    END IF


    sHTML = sHTML & ofv.GenTabla("","aaa","","50% ","","0","","0","0","","",Formulario,"si")


  'Botonera
  Agregar = ""

  Buscar = ""
  
  sAccion = "GLOSMANAGERDSC.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Buscar,Celda,Agregar)

  


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>