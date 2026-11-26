<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sOpcion = request.querystring("OPC")
  if sOpcion <> "" then sOpcion = Ucase(sOpcion) 
  sconte = request.querystring("OConte")
  if sconte = "" then sconte = "NO"
  sconteK = request.querystring("OConteK")
  sconteD = request.querystring("OConteD")

  CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  CantRegAMover2 = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  CantRegAMostrar2 = 1


'Tipobjeto=BLK&Codobjeto=285&Menu=x&Sector=x 

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover) & "_Oconte=" & sconte  & "_OconteK=" & sconteK
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2) & "_Oconte=" & sconte  & "_OconteK=" & sconteK
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2) & "_Oconte=" & sconte  & "_OconteK=" & sconteK

sHTML = "<html><meta HTTP-EQUIV='expires' CONTENT='0' ><META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=windows-1252'>" 
'sHTML = sHTML & "<script language=JavaScript for=window event=onload >" & chr(13)
'sHTML = sHTML & "window.moveTo(0,0);" & chr(13)
'if ofv.explorador = "MSIE" then
'sHTML = sHTML & "window.document.title =" & chr(34) & "COR EManager - Edicion de Esquemas Normativos" & chr(34) & ";" & chr(13)
'end if
'sHTML = sHTML & "</script>" & chr(13)

sHTML = sHTML & "<frameset rows=150,* FRAMESPACING=0  FRAMEBORDER=0 BORDER=0 >"


sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "EdTablasHead.asp?Tipobjeto="
sHTML = sHTML & sTipo & "&Codobjeto=" & nCodigo  & chr(34)
sHTML = sHTML & " NAME=tabhead SCROLLING=auto>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=tabcont "
sHTML = sHTML & "src=" & chr(34) & "EdTablasCont.asp?Tipobjeto=" & sTipo
sHTML = sHTML & "&Codobjeto=" & nCodigo & chr(34)
sHTML = sHTML & " SCROLLING=auto></frameset>"
sHTML = sHTML & "<body  ><head><title>COR EManager - Edicion de tablas</title></head></body></html>"
 
end if
 
response.write sHTML

%>