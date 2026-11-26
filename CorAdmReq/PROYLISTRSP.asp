<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sTIPO = request.querystring("TIPO")
  sopc = request.querystring("opc")

  IF SOPC = "" THEN
     SOPC = "NO"
  END IF

  IF sTIPO = "" THEN
     sTIPO = "PRY"
  END IF

  sRESP = "R"


session("ModMain") = "LR"


  RESPONSE.REDIRECT "PROYLISTDSP.ASP?OPC=" & SOPC & "&TIPO=" & STIPO & "&RESP=" & SRESP





end if


response.write sHTML


%>