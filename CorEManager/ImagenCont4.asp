<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("ID")
    sNIV1   = request.querystring("NIV1")
    sNIV2   = request.querystring("NIV2")
     sNIV3   = request.querystring("NIV3")
     sNIV4   = request.querystring("NIV4")
    sNIV5   = request.querystring("NIV5")

  set cn = ofv.conectar(ofv.strconn2)
  aplcheck()
  strsql = "Select * From Comparc Where codcomp = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  i = 0
  redim preserve aImagen(i)
  redim preserve aDescri(i)
  redim preserve aSecuen(i)
  do until rs.eof
    redim preserve aImagen(i)
    redim preserve aDescri(i)
    redim preserve aSecuen(i)
    aImagen(i) = rs(3)
    if rs(4) <> "" then
       aDescri(i) = "Descripcion:&nbsp;" & rs(4)
    else
       aDescri(i) = "Archivo:&nbsp;" & rs(3)
    end if
    if rs(5) <> "" then
       aSecuen(i) = "Cuerpo:&nbsp;" & rs(5)
    else
       aSecuen(i) = "&nbsp;"
    end if
    rs.movenext
    i = i + 1
  loop

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)




  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")

  for i = lbound(aImagen) to ubound(aImagen)
      sHTML = sHTML & "<tr>"
 '   sHTML = sHTML & "<p >"
    if aImagen(i) = "" then
      sHTML = sHTML & "<td  class=ttt   >"
      sInput = "No poseee ARCHIVOS asociados"
      sImagen = ""
      sHTML = sHTML & sInput
      sHTML = sHTML & "</td>"
    else
      sHTML = sHTML & "<td  class=uts width=5% >"
      sImagen = ofv.GenImage("","","","","Iconos\ICorTema.jpg","1","")
      sHref =  "ImagenCont5.asp?Archivo=" & aImagen(i)
      sLink = ofv.GenLink("","",sHref,"_blank","","",sImagen,"si")
      sHTML = sHTML & sLink
      sHTML = sHTML & "</td>"
      sHTML = sHTML & "<td  class=uti >"
      sInput = aDescri(i) & aSecuen(i)
      sHref =  "ImagenCont5.asp?Archivo=" & aImagen(i)
      sLink = ofv.GenLink("","nt",sHref,"_blank","","",sInput,"si")
      sHTML = sHTML & sLink
      sHTML = sHTML & "</td>"
    end if
'    sHTML = sHTML  & sLink & "</p>"
    sHTML = sHTML & "</tr>"
  next
  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</font>"

    '----------------------------------------------
        '''''''''''''''''''''''''''''''''''''''''''''''
        '' ACTIVACION PUNTERO DEL LINK
        ''
        ''''''''''''''''''''''''''''''''''''''''''''''''
         SPLFUNC=""

        IF session("APLMDCLI")="BC" THEN
         IF SNIV1 = "6" AND SNIV2 <> "0" THEN

            'SDOCC =  SNIV2
          set cnf = ofv.conectar(ofv.strconn2)
    	  Strsql = "SELECT archivo FROM comparc WHERE OBJID = 0" & SNIV2
          'response.write strsql
    				SET Xrs = ofv.crearconsultaEx(StrSql, cnf, 1, parametros)
    				IF NOT Xrs.EOF THEN
    				   'sdirec = Xrs(1) & ""

                      'if sdirec = "" then sdirec = "Archivos"
                       sarch =  Xrs(0)

                    SPLFUNC="<form target=_new name='dispara' method=post action='ImagenCont5.asp?Archivo=" &sarch &"'></form><script language='javascript' >document.getElementById('dispara').submit();</script>"
    				END IF
                CALL ofv.cerrarconsulta(Xrs)
                call ofv.cerracoonn( cnf )

         ELSE

            SPLFUNC = ""

         END IF
       END IF
     sHTML = sHTML &  SPLFUNC
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>