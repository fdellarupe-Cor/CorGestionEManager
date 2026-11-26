<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTSECTORES.inc" -->
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

  sHTML = ofv.FormHeader(session("FormN"))
  sMenu = Session("menu")

  set cn = ofv.conectar(ofv.strconn2)
  Cobjun=1
  strsql = "Select Version,estado,objetos.objun From Objprty "
  strsql = strsql & " inner join objetos on(objetos.codobjeto=Objprty.codobjeto) "
  strsql=strsql & " Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
     Cobjun=rs(2)
    if rs(0) = "V" then bVersion = "SI"
   Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)


shtml = shtml & CorPnlFnc("N","")



 sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  Celda = ofv.GenCelda("","ut","","100% colspan=2","","","","&nbsp;","si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")



  styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:90%;overflow:auto;" & styl & "' >"


  sHTML = sHTML & GENFMTSECTORES(nCodigo,sTipo)

  sHTML = sHTML & "</div>"
  sHTML = sHTML & "</div>"

  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""

   if bVersion = "SI" then
      if SESSION("VIEWMOD") <> "N" then

       '  xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Celda   = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp    = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML   = sHTML & ConsCambios(sasp)


      end if


   end if


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")


  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML  & GTCOBS() & "</body></html>"
end if

response.write sHTML

%>