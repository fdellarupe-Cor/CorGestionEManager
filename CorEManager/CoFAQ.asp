<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTFAQ.inc" -->
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

  set cn = ofv.conectar(ofv.strconn2)

     shtml = ofv.FormHeader(session("FormN")) 


 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"


  sHTML = sHTML &  "<div id='mt2' align=center  style='display:none;height:350px;overflow:auto;" & styl & "' >"


     shtml = shtml & GENFMTFAQ(nCodigo,STIPO)


'''''''''''''''''''''''''''''
'''''''
  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  'sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")



     shtml = shtml & "</body></html>"


end if

response.write sHTML

%>