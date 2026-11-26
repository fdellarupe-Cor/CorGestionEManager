<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sTipo = request.querystring("Tipo")
  sNombre = request.querystring("Nombre")
  vuelta = ofv.vueltaasp2("")

  sHTML = ofv.FormHeader(session("FormN"))

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Tipocomp Where tipo='" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    simage = rs(2)
    sPerfil = rs(4)
  end if
  call ofv.cerrarconsulta(rs)

  Strsql = "Select * From Complementos Where tipocomp='" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    'Eliminar
    sHref = "asp/EliminarComplemento.asp?ID=" & ofv.convertircar(rs(0)," ","_")
    sHref = sHref & "&Tipo=" & sTipo & "&Volver=" & vuelta
    sHref = sHref & "&Eliminar=NSNC"
    Celda = ofv.BotonEliminar(rs(3),sHref)

    sHref = "MenuComplementos.asp?Codcomp=" & rs(0) & "&Tipocomp=" & rs(2) & "&Nombre="
    sHref = sHref & sNombre & "&Complemento=" & ofv.convertircar(rs(3)," ","_") & "&Perfil="
    sHref = sHref & sPerfil
    sImagen = ofv.GenImage("","","","","Iconos\" & simage,"0","")
    sLink = ofv.GenLink("","nt",sHref,"Master","","",sImagen,"si")
    Celda = Celda & ofv.GenCelda("","","","","","","",sLink,"si")

    sHref = "MenuComplementos.asp?Codcomp=" & rs(0) & "&Tipocomp=" & rs(2) & "&Nombre="
    sHref = sHref & sNombre & "&Complemento=" & ofv.convertircar(rs(3)," ","_") & "&Perfil="
    sHref = sHref & sPerfil
    sLink = ofv.GenLink("","nt",sHref,"Master","","",rs(3),"si")
    Celda = Celda & ofv.GenCelda("","","","100%","","","",sLink,"si")

    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    rs.movenext
  loop	
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si")

  'Agregar
  sInput = ofv.GenerarInput("","Agregar","submit","","10","bt","")
  Celda = ofv.GenCelda("","","","","","","",sInput,"si")
  sAccion = "Agregar2.asp?Tipo=" & sTipo & "&Nombre=" & sNombre & "&Volver=" & vuelta
  Agregar = ofv.GenForm("Agregar","",sAccion,"","post","",Celda,"si")
  sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,"no","")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>