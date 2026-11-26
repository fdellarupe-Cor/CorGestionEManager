<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTARCHIVOS.inc" -->
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
  sdocum = request.querystring("sdocum")
  sKIND = request.querystring("SKIND")

  IF SKIND <> "" THEN
     SESSION("SKIND") = SKIND
  ELSE
     SKIND = SESSION("SKIND")
  END IF 


  ''   SKIND = SESSION("SKIND")

  sft   = "&nbsp;<b>"
  sfta  = "&nbsp;<b>"
  sfidc = "<font face=corflaticon size=2 >0</font>" 
  sfida = "<font face=corflaticon size=2 >1</font>" 
  sfitc = "<font face=corflaticon size=2 >0</font>" 
  sfita = "<font face=corflaticon size=2 >1</font>" 
  sftc  = "</b>" 
  
  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_sdocum=" & sdocum

  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)


  set cn = ofv.conectar(ofv.strconn2)


  sHTML   = sHTML & GENFMTARCHIVOS(nCodigo,STIPO)

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & GTCOBS()


  sHTML = sHTML & "</body></html>"
  response.write sHTML


  call ofv.cerrarconn(cn)




end if



%>