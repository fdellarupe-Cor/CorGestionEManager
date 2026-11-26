<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  TA = request.querystring("TA")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sCod = request.querystring("Cod")
  Volver = request.querystring("Volver")
  Volver = Volver & "?Cod=" & sCod
  sxopc1 = ""

     sHTML = ofv.FormHeader("Agregar Usuario")
    '----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"


  'SHTML = SHTML &  GenBuscaUsrEx("S","S","N","S","N","N","N","N","N","S","","","","","","","","","","") 'GenBuscaUsr("S","S","N","S",CHR(9),CHR(9),CHR(9),CHR(9))

    IF session("APLMDCLI")="BPN" THEN
      wwGenBuscaUsr= GenBuscaUsrEx("S","S","N","S","N","N","N","N","N","S","","","","","","","","","","")
    ELSE
      wwGenBuscaUsr= GenBuscaUsrEx("S","S","N","S","N","N","N","N","N","N","","","","","","","","","","")
    END IF

     SHTML = SHTML &  wwGenBuscaUsr

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)

        strsql = "Select * From objDist "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

        f1a = "asp/AgregarEN.asp?Tabla=OBJDist"
        f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
        f1a = f1a & "&Tipobjeto=" & sTipo

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")



        if SESSION("FBSLC") = "U" then

           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(4).NAME & "' >"

        elseif SESSION("FBSLC") = "E" then

           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(5).NAME & "' >"
           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='TODOS' NAME='" & RS(4).NAME & "' >"
        elseIF SESSION("FBSLC")="PORGRUPOEXT" THEN
        ' sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(5).NAME & "' >"
		 sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='GRM-" & SESSION("FBUSR")& "' NAME='" & RS(4).NAME & "' >"


        else
          ' RESPONSE.WRITE SESSION("FBSLC")

           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(5).NAME & "' >"
           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='SECT-' NAME='" & RS(4).NAME & "' >"


        end if




      ELSE

        sAccion = "edobjDistDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF


  if TA = "xU" then



  ELSEif TA = "XU" then

  set cn = ofv.conectar(ofv.strconn0)

  strsql = "Select * From Usuarios  order by descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  do until rs.eof

     sxopc1 = sxopc1 & "<option value=" & rs(0) & ">Usuario:" & rs(1) & "</option> "

     rs.movenext
  loop

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From objDist "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  sHTML = ofv.FormHeader("Agregar Usuario")

  f1a = "asp/AgregarEN.asp?Tabla=OBJDist"
  f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
  f1a = f1a & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Usuario","si")
  h1 = h1 & ofv.GenCelda("","","","2%","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  t1 = ofv.GenCelda("","","","","","","","<select class=at name=" & rs(4).name & " >" & sxopc1 & "</select>","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 ,"si")
  sHTML = sHTML & ofv.GenFTabla()

  elseif TA = "xE" then

 ' TA = "E" - (Agrega empresa completa)

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From clienteplan  order by cliente"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  do until rs.eof


        sxopc1 = sxopc1 & "<option value=" & rs(0) & ">" & rs(1) & "</option> "


  rs.movenext
  loop

  call ofv.cerrarconsulta(rs)


  strsql = "Select * From objDist "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  sHTML = ofv.FormHeader("Agregar Empresa")

    f1a = "asp/AgregarEN.asp?Tabla=OBJDist"
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Empresa","si")
  h1 = h1 & ofv.GenCelda("","","","2%","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  t1 = ofv.GenCelda("","","","","","","","<select class=at name=" & rs(5).name & " >" & sxopc1 & "</select>","si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(4).name,"TODOS","hidden","","","",""),"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 ,"si")
  sHTML = sHTML & ofv.GenFTabla()

  elseif ta = "xx" then

 ' TA = "S" - (Agrega un sector completo)

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From Eco  order by descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  do until rs.eof


        sxopc1 = sxopc1 & "<option value=" & rs(0) & ">" & rs(1) & "</option> "


  rs.movenext
  loop

  call ofv.cerrarconsulta(rs)


  strsql = "Select * From objDist "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  sHTML = ofv.FormHeader("Agregar Sector")

    f1a = "asp/AgregarEN.asp?Tabla=OBJDist"
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Usuario","si")
  h1 = h1 & ofv.GenCelda("","","","2%","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  t1 = ofv.GenCelda("","","","","","","","<select class=at name=" & rs(5).name & " >" & sxopc1 & "</select>","si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(4).name,"SECT-","hidden","","","",""),"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 ,"si")
  sHTML = sHTML & ofv.GenFTabla()

  end if


  'Botones Aceptar y Cancelar
  sAccion = "edobjDistDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
   '--------------------------------------------------------------------------------
     sHTML = sHTML & "</span></div>"

     Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

     'Estructura del Indice

     sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "<script language='javascript' > "
       IF REQUEST.FORM("XGUMASTER")="" THEN
         sHTML = sHTML & " clactfrm('U','T','','NOMBRE','NO');document.XGUFRM.submit(); "
       END IF
     sHTML = sHTML & "</script> "
  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>