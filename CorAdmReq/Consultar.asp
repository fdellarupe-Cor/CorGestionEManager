<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

  set cn = ofv.conectar(ofv.strconn5)

  titc = 0

  CODIGO = REQUEST.QUERYSTRING("CODIGO")

  IF CODIGO <> "" THEN  session("NCliente") = CODIGO

  if session("ModMain") = "WK" then
     SEGCKTPSC = REQUEST.QUERYSTRING("SEGCKTPSC")
     IF SEGCKTPSC <> "" THEN  session("SEGCKTPSC") = SEGCKTPSC
     strsql = "UPDATE SEGPASO SET RESPASG = '" & Session("Usuario") & "' WHERE SEGCKTPSC = '" & SESSION("SEGCKTPSC") & "' "
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF
     
  session("ModMain2") = "CP"

sHTML = "<html><script language=JavaScript for=window event=onload >" & chr(13)
sHTML = sHTML & "window.moveTo(0,0);" & chr(13)
if ofv.explorador = "MSIE" then
sHTML = sHTML & "window.document.title =" & chr(34) & "COR OnTime - Consulta de Requerimientos" & chr(34) & ";" & chr(13)
end if
sHTML = sHTML & "</script>" & chr(13)


sHTML = sHTML & "<frameset rows=50,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "BarraEN.asp"
sHTML = sHTML & chr(34) & " NAME=Barra SCROLLING=auto style='border-style:solid;border-width:2;border-color:#6699cc;' >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Principal "
if session("ModMain") = "WK" then
   sHTML = sHTML & "src=" & chr(34) & "CoWKFREQUER.asp"& chr(34)
ELSE
   sHTML = sHTML & "src=" & chr(34) & "CoREQUER.asp"& chr(34)
END IF
sHTML = sHTML & " SCROLLING=auto style='border-style:solid;border-width:2;border-color:#6699cc;' ></frameset>"
sHTML = sHTML & "<body><head><title>COR OnTime - Consulta de Requerimientos</title></head></body></html>"

response.write sHTML

%>