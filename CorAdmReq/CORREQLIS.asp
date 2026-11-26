<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  sopcion = request.querystring("opcion")
  sopc = request.querystring("opc")

  IF SOPC <> "" THEN
     SESSION("ROPC") = SOPC
  ELSE
     SOPC = SESSION("ROPC")
  END IF

session("ModMain") = "LS"

  IF sTIPO = "" THEN
     sTIPO = "REQ"
  END IF


 LISTA = ""
 DESTINO = ""


  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


'' RESPONSE.WRITE session("Form") & " " & bAgregar & " "  & bModificar & " "  & bSuprimir & " "  & bConsultar & " " 

 SELECT CASE SOPC
        CASE "A"
             RESPONSE.REDIRECT "PROYLISTDSP.ASP?OPC=NO&TIPO=" & STIPO & "&RESP=NO"

        CASE "R"
             RESPONSE.REDIRECT "PROYLISTDSP.ASP?OPC=NO&TIPO=" & STIPO & "&RESP=R"

        CASE "O"
             RESPONSE.REDIRECT "PROYLISTDSP.ASP?OPC=NO&TIPO=" & STIPO & "&RESP=O"

        CASE "L"
          '''   RESPONSE.REDIRECT "PROYLISTDSP.ASP?OPC=NO&TIPO=" & STIPO & "&RESP=A&SEL=A"
             RESPONSE.REDIRECT "PROYREQTREE.ASP?XSET=XGS&XMOD=" & STIPO & "&XDST=ORG"

        CASE "S"
             RESPONSE.REDIRECT "PROYLISTDSP.ASP?OPC=NO&TIPO=" & STIPO & "&RESP=A&SEL=C"

 END SELECT


end if


response.write sHTML


%>