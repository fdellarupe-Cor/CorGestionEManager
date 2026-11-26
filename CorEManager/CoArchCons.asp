<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTARCHIVOS.inc" -->
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
  sdocum = request.querystring("sdocum")
  sKIND = request.querystring("SKIND")
  sNIV1   = request.querystring("NIV1")
  sNIV2   = request.querystring("NIV2")
   sNIV3   = request.querystring("NIV3")
   sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")


  IF SKIND <> "" THEN
     SESSION("SKIND") = SKIND
  ELSE
     SKIND = SESSION("SKIND")
  END IF

  ''   SKIND = SESSION("SKIND")

  sft   = "&nbsp;<font>"
  sfta  = "&nbsp;<font>"
  'sfidc = "<font face=corflaticon size=2 >0</font>"
  sfidc = "<i class='las la-folder'></i>"
  sfida = "<font face=corflaticon size=2 >1</font>"
  sfitc = "<font face=corflaticon size=2 >0</font>"
  sfita = "<font face=corflaticon size=2 >1</font>"
  sftc  = "</font>"

  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_sdocum=" & sdocum

  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------

  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  set cn = ofv.conectar(ofv.strconn2)
  If SESSION("APLMDCLI") = "BV" Or SESSION("APLMDCLI") = "BGAL"  Or SESSION("APLMDCLI") = "BNX" Or SESSION("APLMDCLI") = "CORF" Or SESSION("APLMDCLI") = "COR"  Or SESSION("APLMDCLI") = "BC"  Or SESSION("APLMDCLI") = "BP" Or SESSION("APLMDCLI") = "BPN" Then

   Set cnbsq = OFV.conectar(strconn8)
  end if
  strsqlx = "Select Version From Objprty Where Objprty.codobjeto =0" & nCodigo

     set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
     bVersion = "NO"
     if not rs.eof then
       if rs(0) = "V" then bVersion = "SI"
     end if
   call ofv.cerrarconsulta(rs)


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:90%;overflow:auto;" & styl & "' ><span id=111>"


 ' sHTML   = sHTML & GENFMTARCHIVOS(nCodigo,STIPO)

   IF session("APLMDCLI")="BC"  THEN

     '----------------------------------------------
         '''''''''''''''''''''''''''''''''''''''''''''''
         '' ACTIVACION PUNTERO DEL LINK
         ''
         ''''''''''''''''''''''''''''''''''''''''''''''''
          SPLFUNC=""
          IF SNIV1 = "10" AND SNIV2 <> "0" THEN

             'SDOCC =  SNIV2
     	  Strsql = "SELECT archivo, directorio  FROM Objhtm WHERE OBJID = 0" & SNIV2

     				SET Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
     				IF NOT Xrs.EOF THEN
     				   sdirec = Xrs(1) & ""
                      ' response.write "222"
                       if sdirec = "" then sdirec = "Archivos"
                        sarch = sdirec & "/" & Xrs(0)
                     IF session("APLMDCLI")="BC"  THEN
                      IF INSTR(sNIV5,"PORTADA") > 0 THEN
                      SPLFUNC="<form target=_new name='dispara' id='dispara' method=post action='CoArchivos3.asp?Archivo=" &sarch &"'></form><script language='javascript' >document.getElementById('dispara').submit();</script>"
                      ELSE
                      SPLFUNC="<form target=_top name='dispara' id='dispara' method=post action='CoArchivos3.asp?Archivo=" &sarch &"'></form><script language='javascript' >document.getElementById('dispara').submit();</script>"
                      END IF
                     else
                      SPLFUNC="<form target=_new name='dispara' id='dispara' method=post action='CoArchivos3.asp?Archivo=" &sarch &"'></form><script language='javascript' >document.getElementById('dispara').submit();</script>"
                     end if
     				END IF
                 CALL ofv.cerrarconsulta(Xrs)

          ELSE

             SPLFUNC = ""

          END IF
        if SPLFUNC="" and session("APLMDCLI")="BC" then sHTML = sHTML & GENFMTARCHIVOS(nCodigo,STIPO)
       sHTML = sHTML &  SPLFUNC

   else
   sHTML   = sHTML & GENFMTARCHIVOS(nCodigo,STIPO)
   end if





  sHTML = sHTML & "</span></div>"

  sHTML = sHTML & "</div>"


  Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
  Filann = ofv.GenRow("","","","","","","",Celda,"si")

  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%   ","","0","","1","0","","",Filann,"si")

  'Estructura del Indice




    sHTML = sHTML & CorPnlIndx("N","")



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
  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & GTCOBS()
  sHTML = sHTML & "</body></html>"



  call ofv.cerrarconn(cn)




end if

 response.write sHTML

%>