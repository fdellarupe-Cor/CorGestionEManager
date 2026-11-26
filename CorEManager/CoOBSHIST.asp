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
  sTipo   = request.querystring("Tipobjeto")
  sTabla  = request.querystring("Tabla")
  sSubTit  = request.querystring("SubTit")
  filtro=request.querystring("filtro")
  PRN=request.querystring("PRN")
  act=request.querystring("act")

  sft   = "&nbsp;<b>"
  sfta  = "&nbsp;<b>"
  sfidc = "<font face=corflaticon size=2 >0</font>"
  sfida = "<font face=corflaticon size=2 >1</font>"
  sfitc = "<font face=corflaticon size=2 >0</font>"
  sfita = "<font face=corflaticon size=2 >1</font>"
  sftc  = "</b>"




  'if dato <> "N" then
IF PRN  <> "PRINT" THEN
'----------------------------------------------
  sHTML = ofv.FormHeader(session("FormN"))
  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
 IF FILTRO <> "USR" and session("ModMain")="ED" THEN
   	sAccion = "COOBSHIST.asp?PRN=PRINT&FILTRO="&FILTRO&"&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
    sHref   = sAccion
      if BrowserName() = "IE" Then
      Whtml = "<table  width=100% id=xprn ><tr><td width=95% >&nbsp;</td><td width=5% id='mxbtrecx' style='xdisplay:block'  align=center class='bietx'  ><a href='#' onclick="  & chr(34) & "window.open ('" & saccion & "');" &  chr(34) & "  class='bietx' >Imprimir&nbsp;Observaciones</a></td><tr></table>" '& session("ModMain")
      else
      Whtml = "<table  width=100% id=xprn ><tr><td width=95% >&nbsp;</td><td width=5% id='mxbtrecx' style='xdisplay:block'  align=center class='bietx'  ><a href='#' onclick="  & chr(34) & "window.open ('" & saccion & "',window,'Width=800;Height=600;menubar=yes;help:no;resizable:yes;dialogTop:00px;dialogLeft:000px;status:no'); window.focus(); " &  chr(34) & "  class='bietx' >Imprimir&nbsp;Observaciones</a></td><tr></table>" '& session("ModMain")
      end if
    sHTML = sHTML &  Whtml
 END IF
ELSE

        response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
       response.write "<tr>"
       response.write "<td width=100% >"
       response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"800")
       response.write "<table id=margen1 width=800 marginwidth=0  cellspacing=1 cellpadding=2>"
       response.write "<tr>"
       response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
       response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
       response.write "</tr></table>"
'------------------------------------------------
end if


 DIM STABLA(20)

  set cn = ofv.conectar(ofv.strconn2)


 if stipo = "RSK" then

        stabla (1) = "CIRCUITO DE AUTOEVALUACION"


 else

  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sEsquema = rs(1)
     sPerfil = rs(3)
  end if
 '      response.write strsql


  Strsql = "Select * From Objperfildoc Where codobjeto =" & NCodigo & " and tipobjeto ='" & stipo & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     perfilsp = "N"
  end if
  '    response.write strsql

  if perfilsp = "N" then
     call ofv.cerrarconsulta(rs)
     Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  '    response.write strsql
  end if




    for xx = 1 to 15
        stabla (xx) = rs(xx + 16)
    next

 end if




  'strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"

  '   response.write strsql

 'busco el circuito


  if session("ModMain")= "CLz" OR session("ModMain")= "WK"  OR session("ModMain")= "ED" then
      if filtro="USR" THEN
        strsql = " SELECT distinct a.segckt,a.id,a.comienzo from SEGCIRCUITO a "
	    strsql =strsql &  "  inner join segpaso b on (a.segckt=b.segckt)"
	    strsql =strsql &  "  where a.objcod = '"&nCodigo&"' and a.objtipo='"& sTipo & "'  "
	    strsql =strsql &  "  and b.responsable ='"& session("usuario") &"'   and "
	    strsql =strsql &  "  a.model='WKF00016631121722600921928999'"
 	    strsql =strsql &  "   order by a.id desc"
     ELSE
      strsql = "SELECT * from SEGCIRCUITO where model='WKF00016631121722600921928999' and objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
     END IF
  else
     strsql = "SELECT top 1 * from SEGCIRCUITO where model='WKF00016631121722600921928999' and objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
  end if
  'response.write strsql

   set cnx = ofv.conectar(ofv.strconn5)

   set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)

  if  rss.eof then
    shtml= shtml & "<script language=JavaScript>var et=document.getElementById('xprn'); if (et  != null)  {et.style.display="&chr(34)&"none"&chr(34)&";}</script>"
  end if

  ' call ofv.cerrarconsulta(rss)
   'call ofv.cerrarconn(cnx)

    do while   not rss.eof

       idcircuito=rss("segckt")

       IF wasotxt = "S" THEN

          stylob = "UT valign=middle align=center style='border-style:solid;border-width:1;border-color:#990000;' "

      end if



     sImageOK = "<font face='corflaticon' color='#008800' size=3 >&#061;</font>"
     sImagepn = "<font face='corflaticon' color='#dd0000' size=3 >&#061;</font>"



    Fila = ""

    wkproccod = ""

    sFORMhed=""

      strsql = "SELECT * from OBJOBSCMT where  TIPO='OBS' AND "
	  strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'  "
	  strsql = strsql & "  and codcircuito='" & idcircuito & "'  "
	  if filtro="USR" THEN
	    strsql = strsql & " and codusr='" &session("usuario")&"'"
	  end if

	  strsql = strsql & " order by ACTPEND , codcircuito, CODSOL, codusr,fecha desc,hora desc"
'response.write strsql
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if filtro="USR" then
          if not rs.eof then
               CELDAD = ""
               CeldaD = CeldaD & ofv.GenCelda("","utI BGCOLOR='#ddddff' style='border-style:solid;border-color:#888888;border-width:1;' ","","100%","","","","<b>HISTORIAL DE COMENTARIOS DEL DOCUMENTO COMIENZO : " & rss("COMIENZO") &"</b>","si")
               FilaD     = ofv.genrow("","","","","25","","",CeldaD,"si")
               sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","0","","",FilaD,"si")
         end if

    else
     CELDAD = ""
	 CeldaD = CeldaD & ofv.GenCelda("","utI BGCOLOR='#ddddff' style='border-style:solid;border-color:#888888;border-width:1;' ","","100%","","","","<b>HISTORIAL DE COMENTARIOS DEL DOCUMENTO COMIENZO : " & rss("COMIENZO")&"</b>","si")
	 FilaD     = ofv.genrow("","","","","25","","",CeldaD,"si")
     sFORMhed = ofv.GenTabla("","TWIND","","98%","","0","","0","0","","",FilaD,"si")
    end if

   shtml = shtml & sFORMhed


''  Nuevo 21022014
    strsql = "select * from csrelpathdescriex where  "
    strsql = strsql & " cobj = " & nCodigo & " and tobj = '"& sTipo &"'  "
''  where cobj = 1812 and tobj = 'MN'
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'  response.write strsql
    if not rs2.eof then
       sexpr7  = rs2(7)
       stipod  = rs2(19)
       stemad  = rs2(22)
       sstemad = rs2(23)
    end if
''  Nuevo 21022014 Fin

  WKCODSOL = 0




  if not rs.eof then

  ''  Nuevo 21022014
        Celda = ofv.GenCelda("","ut  ","","2%","","","","&nbsp;","si")
        Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")
        Celda = ofv.Gencelda("","VtI colspan=4 BGCOLOR='#efefef' style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","","ORIGEN:&nbsp;" & stemad & " / " & sstemad & " / " & stipod & " " & sexpr7 ,"si")
        Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")
''  Nuevo 21022014 Fin

  end if

  if not rs.eof then

  iform=0

  do while  not rs.eof

     IF WKCODSOL <> RS(3) THEN
        WKCODSOL = RS(3)
  ''    response.write " WKCODSOL " & RS(3)
        Celda = ofv.GenCelda("","ut  ","","2%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","ut  ","","24%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","ut  ","","38%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","ut  ","","38%","","","","&nbsp;","si")
        Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

        Celda = ofv.GenCelda("","VtI colspan=4 BGCOLOR='#ddddff' style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","","OBJETO:&nbsp;" & stabla (WKCODSOL),"si")
        Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")

''  Nuevo 21022014
''      Celda = ofv.GenCelda("","ut  ","","2%","","","","&nbsp;","si")
''      Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")
''      Celda = ofv.Gencelda("","VtI colspan=4 BGCOLOR='#efefef' style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","","ORIGEN:&nbsp;" & stemad & " / " & sstemad & " / " & stipod & " " & sexpr7 ,"si")
''      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")
''  Nuevo 21022014 Fin



     END IF


          SOBJCOMEN = ""
          XGUFBTAB1 = ""

     IF stipo = "RSK" THEN

        IF WKCODSOL2 <> RS(5) THEN
           WKCODSOL2 = RS(5)

           Celda = ofv.GenCelda("","ut colspan=3 ","","","","","","&nbsp;","si")
           Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")





           strsql = ""
           strsql = strsql & " SELECT A.TIPOACCION,A.DESCRIACC,A.SECUEN,B.CODVINOBJ "
           strsql = strsql & " FROM OBJACCION A "
           strsql = strsql & " INNER JOIN OBJETOS B ON (A.CODOBJETO = B.CODOBJETO)  "
           strsql = strsql & " WHERE A.OBJID = 0" & RS(5) & " "

           set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsX1.eof then
              sACTDATA  = rsX1(0)
              sACTDATAS = rsX1(2)
              sACTDATA2 = rsX1(2) & "." & rsX1(1)
              sACTDATA3 = rsX1(3)

              IF sACTDATA3 = "CROBJ" OR sACTDATA3 = "PLOBJ" OR sACTDATA3 = "AIOBJ" THEN

                 SOBJCOMEN = sACTDATA2

              ELSE

                 strsql = ""
                 strsql = strsql & " SELECT DESCRI FROM TIPOACCALT WHERE TIPOACCIONALT = 0" & sACTDATA
                 set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if not rsX1.eof then
                    SOBJCOMEN = sACTDATAS & "." & rsX1(0)
                 END IF

              END IF


           end if

          XGUFBTAB1 = "ACTIVIDAD:&nbsp;" & SOBJCOMEN

        '   Celda = ofv.GenCelda("","VtI colspan=3 BGCOLOR='#F7EEDD' style='border-style:solid;border-color:#000000;border-width:1;' ","","","","","","ACTIVIDAD:&nbsp;" & SOBJCOMEN,"si")
        '   Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")



         END IF



      ELSE



          wtipoL = nCodigo
          wtipOBJ = sTipo

           WKCODSOL2 = RS(4) & ""
           WKCODSOL3 = RS(5) & ""
           WKCODSOL4 = RS(6) & ""

           WNV1 = "" & RS(3)
           WNV2 = "" & RS(4)
           WNV3 = "" & RS(5)
           WNV4 = "" & RS(6)




           if wnv1 = "9" then
           WNV1 = "" & RS(3)
           WNV2 = "" & RS(6)
           WNV3 = "" & RS(4)
           WNV4 = "" & RS(5)

           end if
            tsec=""
           if wnv1 = "11" then
		         strsql = " SELECT tsec FROM CSESTDET "
				  strsql = strsql & " where objid=0" &  WNV3
				'   if marca <> "CA" then
				'     strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
				'  end if
				'  strsql = strsql & " and not ( "
				'  strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
				'  strsql = strsql & " ) and stmodi = '' "
              '    strsql = strsql & " order by TSEC  "
                '  response.write strsql & chr(13)
                  set rsnum = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                  if not rsnum.eof then
                    tsec=  rsnum(0)
                    anum=split(tsec,".")
                    tsec=""
                    delim=""
                    for ai=0 to ubound(anum)
                    tsec=tsec & delim & anum(ai)- 900000
                    delim="."
                    next
                    tsec="<br>Secuencia : " & tsec
                  end if

                  call ofv.cerrarconsulta(rsnum)
           end if




           if wnv1 = "9" or wnv1 = "11" then

           else
              WNV4 = ""
              if wnv1 = "10"  then
                 WNV3 = "0"
              end if
              if wnv1 = "15"  then
                 WNV2 = "0"
                 WNV3 = "0"
              end if


               if wnv1 = "3" then
  			     WNV3 = "0"
			     WNV4 = "0"
               end if



           end if

            XMST = "0"

                XGUFBTDF0 = ""
                XGUFBTDF1 = ""
                XGUFBTDF2 = ""
                XGUFBTDF3 = ""
                XGUFBTDF4 = ""
                XGUFBTAB1 = ""



                   IF WNV2 <> "0" THEN

                      XMST = "3"

                      CALL GETLISTNV2("E",wtipoL,wtipOBJ,WNV1,WNV2)
                      XGUFCOBJKEY = XGUTABLISKEY
                      XGUFCOBJDAT = "- " & XGUTABLISDAT
                      if WKCODSOL <> "111" then
                        XGUFD = "- " & XGUTABLISDAT
                      end if

                      XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                      XGUFBTDF2 = XGUFBTDF2 & " >"

                      XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT   >"
                      XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                      XGUFBTDF2 = XGUFBTDF2 & "</TR>"

                      IF WNV3 <> "0"  and WKCODSOL <> 5 THEN

                         XMST = "4"

                         CALL GETLISTNV3("E",wtipoL,wtipOBJ,WNV1,WNV2,WNV3)
                         XGUFCOBJKEY = XGUTABLISKEY
                         XGUFCOBJDAT = "- " & XGUTABLISDAT
                         'if WKCODSOL <> "111" then
						    XGUFD = XGUFD  & "- " & XGUTABLISDAT
                        ' end if

                         'XGUFD = XGUFD & "<br>- " & XGUTABLISDAT

                         XGUFBTDF3 = XGUFBTDF3 & "<TR  "
                         XGUFBTDF3 = XGUFBTDF3 & " >"

                         XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT   >"
                         XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
                         XGUFBTDF3 = XGUFBTDF3 & "</TR>"
                       '   response.write WNV4 & "<br>"

                         IF WNV4 <> "" THEN

                            XMST = "5"

                            CALL GETLISTNV4("E",wtipoL,wtipOBJ,WNV1,WNV2,WNV3,WNV4)
                            XGUFCOBJKEY = XGUTABLISKEY
                            XGUFCOBJDAT = "- " & XGUTABLISDAT
                            'XGUFD = XGUFD & "<br>- " & XGUTABLISDAT
                           ' if WKCODSOL = "11" then
							   XGUFD =  XGUFD  & " - " & XGUTABLISDAT
							'else
							'  XGUFD = XGUFD &   "- " & XGUTABLISDAT
							'end if


                            XGUFBTDF4 = XGUFBTDF4 & "<TR  "
                            XGUFBTDF4 = XGUFBTDF4 & " >"

                            XGUFBTDF4 = XGUFBTDF4 & "<TD CLASS=UT   >"
                            XGUFBTDF4 = XGUFBTDF4 & XGUFCOBJDAT & "</TD>"
                            XGUFBTDF4 = XGUFBTDF4 & "</TR>"


                         END IF



                      END IF
                   END IF
'response.write XMST & "<br>"

             IF XMST > "2" THEN

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


             END IF

             IF XMST > "3" THEN

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


             END IF

             IF XMST > "4" THEN

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR4""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF4 & "</TABLE>"


             END IF

             XGUFBTAB1 =  XGUFD & tsec

        '   Celda = ofv.GenCelda("","VtI colspan=3 BGCOLOR='#F7EEDD' style='border-style:solid;border-color:#000000;border-width:1;' ","","","","","",XGUFBTAB1,"si")
        '   Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")



      END IF


    Celda = ofv.GenCelda("","ut colspan=3 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

        xxxx = getInfoUser(RS(9))
	if session("SGU_nombre")<> "" then    sInput =  session("SGU_nombre")

   celda = ""

      select case rs("ACTPEND")
	      case 1
	         sInput = "<font face='corflaticon' color='#008800' size=3 >&#061;</font>"
	      case 2
	         sInput = "<font face='corflaticon' color='#dd0000' size=3 >&#061;</font>"
	      case else
	         sInput = "<font face='corflaticon' color='#dd0000' size=3 >&#061;</font>"
      end select

   'if (rs(12) & "") <> "" then
  '    sInput = sImageOK
  ' else
   '   sInput = sImagepn
  ' end if




   Celda =  celda &  ofv.GenCelda("","ut ROWSPAN=2 bgcolor='#efefef' TITLE='Estado de la observacion.' style='border-style:solid;border-color:#888888;border-width:1;' valign='top' ","CENTER","","","","",sInput,"si")

   sInput = RS(9) & "&nbsp;-&nbsp;" & session("SGU_nombre")

   Celda =  celda &  ofv.GenCelda("","uti style='border-style:solid;border-color:#888888;border-width:1;color:#000000;' ","","","","","",sInput,"si")
 ' Fecha y Hora
   sInput = GTFechadorsh(rs(10)) & "&nbsp;-&nbsp;" & GETGLBHMSW(rs(11))

   Celda =  celda &  ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;color:#000000;' ","","","","","",sInput,"si")

   Fila  = fila & ofv.GenRow("","aaa bgcolor='#dddddd' ","","","","","",Celda,"si")


   'Texto


    sInput = trim(rs(13)&"")

    if sinput = "" then sinput = "-----"



''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                    ''
''   Agregado 27022014 - link a documento observado   ''
''                                                    ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          sImagen  = "NO"
          waso1txt = "N"


'''' link relacionados al contenido
          descri   = ""

          wNV1    = rs(3)
          wNV2    = rs(4)
          wNV3    = rs(5)
          wNV4    = rs(6)

 ''       wsobjid   = rs(0)
 ''       wscodob   = rs(1)
 ''       wstipob   = rs(2)
 ''       wsinx3    = rs(6)


 ''       if wNV1 = "9" then
 ''          response.write wNV1    & " wNV1 "
 ''          response.write wNV2    & " wNV2 "
 ''          response.write wNV3    & " wNV3 "
 ''          response.write wsobjid & " wsobjid "
 ''          response.write wscodob & " wscodob "
 ''          response.write wstipob & " wstipob "
 ''          response.write wsinx3  & " wsinx3 "
 ''       end if

          nCodigo = clng(rs("codobjeto"))
          wtipoLc    = tipobjeto

          wUrlDesc = ""

          Strsql = "Select * From csGenObjetos where estado <> 4 and codobjeto = " & nCodigo
          set rst2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rst2.eof then
             wNomL = rst2(1)
          else
             wNomL =""' "S/D"
          end if
      ''  call ofv.cerrarconsulta(rst2)

      ''  response.write nCodigo & " aaaa "
      ''  response.write WNV1 & " bbbb "

      ''  sActcon = "edestrdet.asp?Codigo=" & nCodigo & "&Tipo=" & sTipo

      ''  Agregado 10032014 - Sergio
      select CASE WNV1
      	     CASE "1"   'PROPIEDADES'
      	          sActcon = "edPropiedades.asp?" & "ID="
      	     CASE "2"   'CONSIDERACIONES'
      	          sActcon = "EDNWEDT.asp?" & "ID="
             CASE "3"   'SECTORES'
      	          sActcon = "edSectores.asp?" & "ID="
             CASE "4"   'EVENTOS'
      	          sActcon = "edEventos.asp?" & "ID="
             CASE "5"   'PROCEDIMIENTOS
                  sActcon = "edProcedimientos.asp?" & "SKIND=1" & "&ID="
''                sActcon = "edProcedDET.asp?" & "SKIND=" & SKind & "&ID="
             CASE "6"   'DESCRIPCION
                  sActcon = "edcontroles.asp?" & "SKIND=1" & "&ID="
             CASE "7"   'CONTROLES
                  sActcon = "edcontroles.asp?" & "SKIND=1" & "&ID="
             CASE "8"   'ESQUEMAS
                  sActcon = "edesqueDET.asp?" & "xeven="
             CASE "9"   'TABLAS
''                response.write " WNV4 " & WNV4
''                sActcon = "EdTablasBLKCON.asp?" & "Tabla=" & WNV4 & "&INX3="
                  sActcon = "EdTablasBLK.asp?" & "Tabla=" & WNV4 & "&INX3="
             CASE "10"  'DOCUMENTOS
      	          sActcon = "edarchivos.asp?" & "SKIND=1"  & "&ID="
      	     CASE "11"  'CONTENIDO
      	          sActcon = "edestrdet.asp?" & "ID="
      	     CASE "12"  'CURSOGRAMA
                  sActcon = "edProcedimientos.asp?" & "SKIND=1" & "&ID="
''    	          sActcon = "edProcedDET.asp?" & "SKIND=" & SKind & "&ID="
      	     CASE "13"  'CONTRATOS
      	          sActcon = "edarchivos.asp?" & "SKIND=2" & "&ID="
      	     CASE "14"  'FAQ
      	          sActcon = "edGeneral.asp?" & "ID="
      	     CASE "15"  'GENERAL
      	          sActcon = "edGeneral.asp?" & "ID="
    	end select

    	''  Fin Agregado 10032014 - Sergio

             sActcon = sActcon & WNV2 & "&CIRCUITO=" &IDCIRCUITO & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
             sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
             sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3 & "&histobs=S"

        ''   response.write " sActcon " & sActcon

             sHref    = sActcon
          '  response.write sActcon
           ' sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;" & Chr(34)


          call ofv.cerrarconsulta(rst2)

     ''    Celda = ofv.GenCelda("","ut  ","","2%","","","","&nbsp;","si")
     ''    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")


     ''    Celda = ofv.Gencelda("","VtI colspan=4 BGCOLOR='#efefef' style='border-style:solid;border-color:#000000;border-width:1;' ","","","","","",sLink,"si")
     ''    Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")

          if session("ModMain") <> "CL" AND act <> "N"  THEN
           sLink    = OFV.GenLink("", "nt  ",sHref, "", "", "", XGUFBTAB1  , "si")
          else
          sLink    = OFV.GenLink("", "nt  ","", "", "", "", XGUFBTAB1  , "si")
          end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                    ''
'' fin agregado 27022014 - link a documento observado ''
''                                                    ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   celda = ""
   Celda =  celda &  ofv.GenCelda("","uti  BGCOLOR='#F7EEDD' style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sLink ,"si")
   Celda =  celda &  ofv.GenCelda("","uti COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;color:#000000;' ","","","","","",sInput  ,"si")

   Fila  = fila & ofv.GenRow("","aaa bgcolor='#F0F0F0'  ","","","35","","",Celda,"si")

rs.movenext
iform=iform+1
loop


else
    if sFORMhed <> "" then
      Celda = ofv.GenCelda("","ut colspan=3 ","","","","","","&nbsp;","si")
      Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

      Celda = ofv.GenCelda("","uti colspan=3 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","","No hay observaciones para el elemento.","si")
      Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")
    end if

end if

    Celda = ofv.GenCelda("","ut colspan=3 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

  shtml =shtml & ofv.GenTabla("","aaa  ","","98% ","","0","","1","1","","",fila,"si")
  rss.movenext
loop
call ofv.cerrarconsulta(rss)
call ofv.cerrarconn(cnx)

 if PRN <> "PRINT" then
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

end if
'--------------------------------------------------------------------------------


  sHTML = sHTML & "</body></html>"


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)




end if

  response.write sHTML

%>