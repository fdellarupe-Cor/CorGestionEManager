<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0


  WKF = REQUEST.QUERYSTRING("WKF")

  IF WKF <> "" THEN  session("WKF") = WKF

  set cn = ofv.conectar(ofv.strconn5)

  StrSql = "Select CKTDSC From WKFCKTSEG WHERE WKFCOD = '" & WKF & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     session("WKFDSC") = RS(0)
  END IF
     
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


sHTML = "<html>" & chr(13)
sHTML = sHTML & "<frameset rows=210,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "SEGCKTSTARTHC.asp"
sHTML = sHTML & chr(34) & " NAME=TOPE SCROLLING=auto>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=BUSCA "
sHTML = sHTML & "src=" & chr(34) & "SEGCKTSTARTSC.asp" & chr(34) & " SCROLLING=auto></frameset>"
sHTML = sHTML & "<body><head><title>COR EManager - Edicion de Esquemas Normativos</title></head></body></html>"

response.write sHTML

%>