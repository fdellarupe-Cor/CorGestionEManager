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



  sft   = "&nbsp;<b>"
  sfta  = "&nbsp;<b>"
  sfidc = "<font face=wingdings size=2 >0</font>"
  sfida = "<font face=wingdings size=2 >1</font>"
  sfitc = "<font face=wingdings size=2 >0</font>"
  sfita = "<font face=wingdings size=2 >1</font>"
  sftc  = "</b>"



  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------



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


  Strsql = "Select * From Objperfildoc Where codobjeto =" & NCodigo & " and tipobjeto ='" & stipo & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     perfilsp = "N"
  end if

  if perfilsp = "N" then
     call ofv.cerrarconsulta(rs)
     Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if




    for xx = 1 to 15
        stabla (xx) = rs(xx + 16)
    next

 end if




  strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"

 ' response.write strsql

 'busco el circuito

   set cnx = ofv.conectar(ofv.strconn5)

   set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)

   if not rss.eof then idcircuito=rss("segckt")

   call ofv.cerrarconsulta(rss)
   call ofv.cerrarconn(cnx)

   IF wasotxt = "S" THEN

    stylob = "UT valign=middle align=center style='border-style:solid;border-width:1;border-color:#990000;' "

   end if



     sImageOK = "<font face='webdings' color='#008800' size=3 >&#061;</font>"
     sImagepn = "<font face='webdings' color='#dd0000' size=3 >&#061;</font>"



    Fila = ""

    wkproccod = ""




    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","utI BGCOLOR='#ddddff' style='border-style:solid;border-color:#888888;border-width:1;' ","","100%","","","","HISTORIAL DE COMENTARIOS DEL DOCUMENTO","si")
    FilaD     = ofv.genrow("","","","","25","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","0","","",FilaD,"si")


   shtml = shtml & sFORMhed




  WKCODSOL = 0


  strsql = "SELECT * from OBJOBSCMT where  "
  strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'  "
  strsql = strsql & "  "
  strsql = strsql & " order by ACTPEND , codcircuito, CODSOL, CODINX1, CODINX2, CODINX3, CODINX4, codusr, objid"
 'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then


  iform=0

  do while  not rs.eof

     IF WKCODSOL <> RS(3) THEN
        WKCODSOL = RS(3)

        Celda = ofv.GenCelda("","ut  ","","2%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","ut  ","","24%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","ut  ","","38%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","ut  ","","38%","","","","&nbsp;","si")
        Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")



        Celda = ofv.GenCelda("","VtI colspan=4 BGCOLOR='#ddddff' style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","","OBJETO:&nbsp;" & stabla (WKCODSOL),"si")
       Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")



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


                      XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                      XGUFBTDF2 = XGUFBTDF2 & " >"

                      XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT   >"
                      XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                      XGUFBTDF2 = XGUFBTDF2 & "</TR>"

                      IF WNV3 <> "0"  THEN

                         XMST = "4"

                         CALL GETLISTNV3("E",wtipoL,wtipOBJ,WNV1,WNV2,WNV3)
                         XGUFCOBJKEY = XGUTABLISKEY
                         XGUFCOBJDAT = "- " & XGUTABLISDAT


                         XGUFBTDF3 = XGUFBTDF3 & "<TR  "
                         XGUFBTDF3 = XGUFBTDF3 & " >"

                         XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT   >"
                         XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
                         XGUFBTDF3 = XGUFBTDF3 & "</TR>"


                         IF WNV4 <> "" THEN

                            XMST = "5"

                            CALL GETLISTNV4("E",wtipoL,wtipOBJ,WNV1,WNV2,WNV3,WNV4)
                            XGUFCOBJKEY = XGUTABLISKEY
                            XGUFCOBJDAT = "- " & XGUTABLISDAT


                            XGUFBTDF4 = XGUFBTDF4 & "<TR  "
                            XGUFBTDF4 = XGUFBTDF4 & " >"

                            XGUFBTDF4 = XGUFBTDF4 & "<TD CLASS=UT   >"
                            XGUFBTDF4 = XGUFBTDF4 & XGUFCOBJDAT & "</TD>"
                            XGUFBTDF4 = XGUFBTDF4 & "</TR>"


                         END IF



                      END IF
                   END IF


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
	         sInput = "<font face='webdings' color='#008800' size=3 >&#061;</font>"
	      case 2
	         sInput = "<font face='webdings' color='#dd0000' size=3 >&#061;</font>"
	      case else
	         sInput = "<font face='webdings' color='#dd0000' size=3 >&#061;</font>"
      end select

   'if (rs(12) & "") <> "" then
  '    sInput = sImageOK
  ' else
   '   sInput = sImagepn
  ' end if




   Celda =  celda &  ofv.GenCelda("","ut ROWSPAN=2 bgcolor='#efefef' TITLE='Estado de la observacion.' style='border-style:solid;border-color:#888888;border-width:1;' valign='top' ","CENTER","","","","",sInput,"si")

   sInput = RS(9) & "&nbsp;-&nbsp;" & session("SGU_nombre")

   Celda =  celda &  ofv.GenCelda("","uti style='border-style:solid;border-color:#888888;border-width:1;color:#000000;' ","","","","","",sInput,"si")

   sInput = GTFechadorsh(rs(10)) & "&nbsp;-&nbsp;" & GETGLBHMSW(rs(11))

   Celda =  celda &  ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;color:#000000;' ","","","","","",sInput,"si")



   Fila  = fila & ofv.GenRow("","aaa bgcolor='#dddddd' ","","","","","",Celda,"si")


   'Texto


    sInput = trim(rs(13)&"")

    if sinput = "" then sinput = "-----"




   celda = ""
   Celda =  celda &  ofv.GenCelda("","uti  BGCOLOR='#F7EEDD' style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",XGUFBTAB1 ,"si")
   Celda =  celda &  ofv.GenCelda("","uti COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;color:#000000;' ","","","","","",sInput  ,"si")

   Fila  = fila & ofv.GenRow("","aaa bgcolor='#F0F0F0'  ","","","35","","",Celda,"si")




rs.movenext
iform=iform+1
loop


else

    Celda = ofv.GenCelda("","ut colspan=3 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

    Celda = ofv.GenCelda("","uti colspan=3 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","","No hay observaciones para el elemento.","si")
    Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")


end if

    Celda = ofv.GenCelda("","ut colspan=3 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

  shtml =shtml & ofv.GenTabla("","aaa  ","","98% ","","0","","1","1","","",fila,"si")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------


  sHTML = sHTML & "</body></html>"


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)




end if

  response.write sHTML

%>