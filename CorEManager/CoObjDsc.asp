<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTOBJDSC.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
'  response.buffer = true

  sTitulo = request.querystring("Titulo")
  sTitulo = ofv.convertircar(sTitulo,"_"," ")

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ID = nCodigo

  sHTML = ofv.FormHeader(session("FormN")) '''ofv.FormBody2

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"


  sHTML = sHTML & ""

  set cn = ofv.conectar(ofv.strconn2)
  nID = ID

  sHTML = sHTML & GENFMTOBJDSC(nCodigo,STIPO)

  sHTML = sHTML & "<br>"

'''''''''''''''''''''''''''''
'''''''
  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  ''sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")





  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>