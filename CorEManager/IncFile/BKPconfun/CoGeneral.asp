<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTGENERAL.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sCodigo = nCodigo

  set cn = ofv.conectar(ofv.strconn2)

    sHTML = ofv.FormHeader(session("FormN"))

'''''''''''''''''
      sImagenEM  = "<FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT>"
      sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"

      shtml   = shtml & "<script language=javascript>" & chr(13)
      shtml   = shtml & "function shwOrig(xusr,XUSRN){" & chr(13)
      shtml   = shtml & "document.all.cx1.style.display='block';" & chr(13)
      shtml   = shtml & "}" & chr(13)

            shtml=shtml & "function chequear(){" & chr(13)
            shtml=shtml & "alert(window.top.CorMtopm.name + ' '  + window.top.CorMtopm.top.name);" & chr(13)
            shtml=shtml & "}" & chr(13)

      shtml   = shtml & "</script>" & chr(13)
'----------------------------------------------

      shtml   = shtml & GENFMTGENERAL(nCodigo,STIPO)



  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</center>"

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML



%>