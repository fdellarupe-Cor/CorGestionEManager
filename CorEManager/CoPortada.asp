<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTportada.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  APLCHECK()
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  genlnk=request.querystring("genlnk")

  sHTML = ofv.FormHeader(session("FormN"))
  sMenu = Session("menu")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
   Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)


shtml = shtml & CorPnlFnc("N","")

if session("APLMDCLI") =  "BCR" or session("APLMDCLI")="BC" OR session("APLMDCLI")="BANELCO" THEN
  shtml= replace(shtml,"<body " ,"<BODY ONLOAD='ClosIndice();' ")
END IF


 sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  Celda = ofv.GenCelda("","ut","","100% colspan=2","","","","&nbsp;","si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")



  styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"


  sHTML = sHTML &  "<div id='mt2'  style='display:block;height:90%;overflow:auto;" & styl & "' ><SPAN ID=1111>"


  sHTML = sHTML & GENFMTportada(nCodigo,sTipo)

  sHTML = sHTML & "</div>"
  sHTML = sHTML & "</div>"

  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""




  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")


  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)



  sHTML = sHTML  & "</SPAN></center></body></html>"
end if

response.write sHTML

%>