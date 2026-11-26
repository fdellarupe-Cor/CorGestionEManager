<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

  ID      = request.querystring("ID")
  sNIV1   = request.querystring("NIV1")
  sNIV2   = request.querystring("NIV2")
  sNIV3   = request.querystring("NIV3")
  sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")

  set cn = ofv.conectar(ofv.strconn2)

 Strsql = "Select tipocomp,codiden, descri From complementos Where codcomp = " & id
 set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 if  not rs.eof then
     stipo = rs(0)
     snomb = rs(1) & "-" & rs(2)
 end if

  strsql = "select * from compcompuesto where compmaster=" & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)
   ' response.redirect "imagencont.asp?ID=" & ID

   response.redirect "EditarComp.asp?codcomp=" & ID & "&tipocomp=" & stipo & "&funcion=C" &  "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 & "&Complemento=" & ofv.convertircar(snomb," ","_")

  else
    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)

    response.redirect "EditarComp.asp?codcomp=" & ID & "&tipocomp=" & stipo & "&funcion=C" &  "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 & "&Complemento=" & ofv.convertircar(snomb," ","_")


  end if
end if

Response.Write sHTML

%>