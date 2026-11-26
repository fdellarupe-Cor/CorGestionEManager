<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

  set cn = ofv.conectar(ofv.strconn5)

  titc = 0

  CODIGO = REQUEST.QUERYSTRING("CODIGO")

  tsk = REQUEST.QUERYSTRING("tsk")

  IF CODIGO <> "" THEN  session("NCliente") = CODIGO

  wtsk = "NO"

  IF tsk <> "" THEN  wtsk = tsk

  VPLAN = REQUEST.QUERYSTRING("VPLAN")

  IF VPLAN <> "" THEN  session("VPLAN") = VPLAN
     
  session("ModMain2") = "CP"




sHTML = "<html><script language=JavaScript for=window event=onload >" & chr(13)
sHTML = sHTML & "window.moveTo(0,0);" & chr(13)
if ofv.explorador = "MSIE" then
sHTML = sHTML & "window.document.title =" & chr(34) & "Cor OnTime - Consulta de PROYECTOS" & chr(34) & ";" & chr(13)
end if
sHTML = sHTML & "</script>" & chr(13)


sHTML = sHTML & "<frameset rows=50,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "BarraPRY.asp?tsk=" & wtsk
sHTML = sHTML & chr(34) & " NAME=Barra SCROLLING=auto style='border-style:solid;border-width:2;border-color:#6699cc;' >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Principal "
if session("ModMain") = "WK" and wtsk <> "NO" then
   sHTML = sHTML & "src=" & chr(34) & "proytaskadv.asp?TSK=" & wtsk & chr(34)
ELSE
   sHTML = sHTML & "src=" & chr(34) & "proyprop.asp" & chr(34)
END IF
sHTML = sHTML & " SCROLLING=auto style='border-style:solid;border-width:2;border-color:#6699cc;' ></frameset>"
sHTML = sHTML & "<body><head><title>Cor OnTime - Consulta de PROYECTOS</title></head></body></html>"

response.write sHTML

%>