<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

sFuncion = request.querystring("funcion")
sACC    = request.querystring("ACC")
sDOCMODAL  = request.querystring("DOCMODAL")
Stipocpt=request.querystring("tipocpt")
IDPEDIDO=REQUEST.QUERYSTRING("CODIGO")

IF sDOCMODAL  = "CL" THEN
   session("ModMain") = "CL"
ELSEIF sDOCMODAL  = "WK" THEN
   session("ModMain") = "WK"
ELSEIF sDOCMODAL  = "WKPED" THEN
   session("ModMain") = "WKPED"
END IF


IF sACC = "" THEN sACC = "NM"

if sFuncion = "" then
   sfUncion = request.querystring & "&Funcion=E"
else
   sfUncion = request.querystring
end if
if Stipocpt="PD" THEN
  session("ncliente")=IDPEDIDO

  IF session("ModMain")="ED" THEN
   response.redirect "EDITARDOC.asp?acc=PF&pedido=VER&DOCMODAL=" & sDOCMODAL
  ELSE
   response.redirect "EDITARDOC.asp?PEDIDO=VER&" & sfUncion
 END IF

ELSEIF Stipocpt="SN" THEN

  session("ncliente")=IDPEDIDO
	segcktpsc = request.querystring("SEGCKTPSC")
  'IF session("ModMain")="ED" THEN
   response.redirect "EDITARDOC.asp?acc=SNN&pedido=VER&DOCMODAL=" & sDOCMODAL & "&segcktpsc=" & segcktpsc
  'ELSE
  ' response.redirect "EDITARDOC.asp?PEDIDO=VER&" & sfUncion
  'END IF

ELSE

response.redirect "EditarDoc.asp?ACC=" & sACC & "&" & sfUncion
END IF
'response.redirect "EDITARDOC.asp?acc=PF&pedido=VER&seg=CLIMN52412655104952828166469"
'RESPONSE.WRITE sfUncion
'response.write "EditarDoc.asp?ACC=" & sACC & "&" & sfUncion


%>