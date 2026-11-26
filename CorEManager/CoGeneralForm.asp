<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTGENERALFORM.inc" -->
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
  sCodigo = nCodigo

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)


    sHTML = ofv.FormHeader(session("FormN")) & corstyler2()
    'sHTML = ofv.FormHeader(session("FormN")) & corstyler2()

'''''''''''''''''
      sImagenEM  = "<FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>&#0252;</FONT> "
      sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>4</FONT></B> "

      shtml   = shtml & "<script language=javascript>" & chr(13)
      shtml   = shtml & "function shwOrig(xusr,XUSRN){" & chr(13)
      shtml   = shtml & "document.all.cx1.style.display='block';" & chr(13)
      shtml   = shtml & "}" & chr(13)

            shtml=shtml & "function chequear(){" & chr(13)
            shtml=shtml & "alert(window.top.CorMtopm.name + ' '  + window.top.CorMtopm.top.name);" & chr(13)
            shtml=shtml & "}" & chr(13)

      shtml   = shtml & "</script>" & chr(13)
'----------------------------------------------

 shtml = shtml & CorPnlFnc("N","")

' if  session("APLMDCLI") =  "BGAL" THEN
 '  shtml= replace(shtml,"<body " ,"<BODY ONLOAD='ClosIndice();' ")
'END IF

 sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  Celda = ofv.GenCelda("","ut","","100% colspan=2","","","","&nbsp;","si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  'sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")



  styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:90%;overflow:auto;" & styl & "' >"




      shtml   = shtml & GENFMTGENERALFORM(nCodigo,STIPO)

  sHTML = sHTML & "</div>"
  sHTML = sHTML & "</div>"


  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</center>"


  Agregar = ""
  Buscar  = ""
  Celda   = ""

   if bVersion = "SI" then
      if SESSION("VIEWMOD") <> "N" then

         'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Celda   = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp    = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML   = sHTML & ConsCambios(sasp)


      end if


   end if



  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")



  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)



  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML



%>