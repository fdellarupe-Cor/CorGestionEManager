<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("ID")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From Complementos Where codcomp = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sNombre = rs(3)
    sCodint = rs(1)
  end if 
  call ofv.cerrarconsulta(rs)

  strsql = "Select archivo From Compimag Where codcomp = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sImagen = rs(0)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = "<html><title>COR E-Manager - Edicion de Complementos</title>"
  sHTML = sHTML & "<body " & ofv.wbody & " ><center><font color=#336699 >"

  if sImagen <> "" then
    'Imagen de Complemento
    sHTML = sHTML & "<h4 ><p align=center>Complemento: " & sCodint & " - " & sNombre
    sHTML = sHTML & "</p></h1><img src=Componentes/" & sImagen & ">"
  else
    sHTML = sHTML & "<h4 ><p align=center >Sin imagen asociada</p></h1>"
  end if
  sHTML = sHTML & "</font></center></body></html>"
end if

Response.Write sHTML

%>