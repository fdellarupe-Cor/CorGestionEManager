<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTOBJDSC.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
'  response.buffer = true

  sTitulo = request.querystring("Titulo")
  sTitulo = ofv.convertircar(sTitulo,"_"," ")

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ID = nCodigo

  sHTML = ofv.FormHeader(session("FormN")) '''ofv.FormBody2

  sHTML = sHTML & "<center><br>"

  set cn = ofv.conectar(ofv.strconn2)
  nID = ID

  sHTML = sHTML & GENFMTOBJDSC(nCodigo,STIPO)


  sHTML = sHTML & "</center><br>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>