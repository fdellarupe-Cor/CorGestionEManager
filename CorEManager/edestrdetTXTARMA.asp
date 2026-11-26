<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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
  IDTXT = request.querystring("IDTXT")
  vuelve = request.querystring("vuelve")
  vuelve = replace(vuelve,"***","_")
  volver = "_Vuelve=" & replace(vuelve,"_","***")
  addqry = replace(vuelve,"_","&")

  sform="modtxt"

  wposi = 0
  wpossal = 0

  vuelta = "../edestrdetTXT.asp?IDTXT=" & IDTXT & volver


    set cn = ofv.conectar(ofv.strconn2)




  sHTML = "<html><body>"

'-------------------------------------------------

response.write sHTML



XXBSQ1= GenBuscaUsrEx("N","N","N","N","N","N","N","S","N","N","","","","","","","","","","")' &  REQUEST.FORM("XGUITEMCD")


response.write  XXBSQ1 &   "</body><script language='javascript' > window.top.LDFOBJMENU('WK4','EXBSUSR') </script>"


 RESPONSE.WRITE   "</BODY></HTML>"



end if
%>





