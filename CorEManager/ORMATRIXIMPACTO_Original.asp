<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  dim xposind
  xposind=0
  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  SPROCMOD = REQUEST.QUERYSTRING("PROCMOD")

  SPROCESO = REQUEST.QUERYSTRING("PROCESO")
  SFILTRO  = REQUEST.QUERYSTRING("FILTRO")
  SFILSEL  = REQUEST.QUERYSTRING("FILSEL")
  SsubFILTRO  = REQUEST.QUERYSTRING("SUBFILTRO")
  SsubFILTRO2 = REQUEST.QUERYSTRING("SUBFILTRO2")
  SsubFILTRO3 = REQUEST.QUERYSTRING("SUBFILTRO3")
  Speriodo = REQUEST.QUERYSTRING("periodo")


  if Speriodo = "" then Speriodo = session("ORPERIODO")

  'response.write Speriodo & "--"  & session("ORPERIODO")
  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

     if Speriodo = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL') and RGVALEAUTOEV = 'S' "
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"
       ' response.write strsql
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           session("ORPERIODOMAP")  = RS(8)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")
           Speriodomap              = session("ORPERIODOMAP")

        end if

      else

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & Speriodo
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           Speriododsc              = RS(1)
           Speriodomap              = RS(8)
        end if

      end if

  Fila = ""
  xTrColor = "#F9F9F9"


  sImagevMa    = "<font face='arial' size=1 xcolor='#aa0000' style='font-weight:900;background-color:#aa0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MUY&nbsp;ALTO&nbsp;</font>"
  sImageva    = "<font face='arial' size=1 xcolor='#f7aa22' style='font-weight:900;background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ALTO&nbsp;</font>"
  sImagevm    = "<font face='arial' size=1 xcolor='#f7EEDD' style='font-weight:900;background-color:#f7EEDD;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MODERADO&nbsp;</font>"
  sImagevb    = "<font face='arial' size=1 xcolor='#6699CC' style='font-weight:900;background-color:#6699CC;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ESCASO&nbsp;</font>"
  sImagevNUL    = "<font face='arial' size=1 xcolor='#00AA00' style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NULO&nbsp;</font>"

  sImageSI    = "<font face='arial' size=1  style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;SI&nbsp;</font>"
  sImageNO    = "<font face='arial' size=1  style='font-weight:900;background-color:#AA0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NO&nbsp;</font>"

  sImagevn    = "<font face='arial' size=1 xcolor='#dddddd' style='font-weight:900;background-color:#dddddd;color:#ffffff' >----</font>"

  wkstylof = "UT style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' "
  wkstylof2 = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "




  DIM TABPROBABILIDAD(10,5)
  DIM TABIMPACTO(10,5)
  DIM TABRIESGOS(50,6)
  DIM TABRIESGOSRES(50,6)
  DIM TABRANRIESGOS(10,5)

  TTABPROB      = 0
  TTABIMPA      = 0
  TTABRIESGO    = 0
  TTABRIESGORES = 0


  'variable utilizada para saber donde comienzan las posiciones de 



  Strsql = "Select * From ORNIVPROBOCURRENCIA ORDER BY VALOR"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABPROBABILIDAD(POSI,0) = RS(0) 'CODIGO
          TABPROBABILIDAD(POSI,1) = RS(1) 'VALOR
          TABPROBABILIDAD(POSI,2) = RS(2) 'DESCRIPCION
          TABPROBABILIDAD(POSI,3) = RS(3) 'PORCENTAJE
          TABPROBABILIDAD(POSI,5) = RS(6) 'COLOR

          TTABPROB = POSI

        RS.MOVENEXT
     LOOP



  END IF




  Strsql = "Select * From ORNIVIMPACTO ORDER BY VALOR"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1


          TABIMPACTO(POSI,0) = RS(0) 'CODIGO
          TABIMPACTO(POSI,1) = RS(1) 'VALOR
          TABIMPACTO(POSI,2) = RS(2) 'DESCRIPCION
          TABIMPACTO(POSI,3) = RS(3) 'MONTODESDE
          TABIMPACTO(POSI,4) = RS(4) 'MONTOHASTA
          TABIMPACTO(POSI,5) = RS(6) 'COLOR

          TTABIMPA = POSI


        RS.MOVENEXT
     LOOP



  END IF



  Strsql = "Select * From ORRANGORIESGO ORDER BY VALOR"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRANRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRANRIESGOS(POSI,1) = RS(1) 'VALOR
          TABRANRIESGOS(POSI,2) = RS(2) 'DESCRIPCION
          TABRANRIESGOS(POSI,3) = RS(3) 'COLOR
          TABRANRIESGOS(POSI,4) = 0     'CANTIDAD

          TTABRANRIESGO = POSI

        RS.MOVENEXT
     LOOP



  END IF




  Strsql = "Select * From ORNIVRIESGOBRUTO ORDER BY VALOR DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRIESGOS(POSI,1) = RS(1) 'VALOR

          TABRIESGOS(POSI,3) = RS(3) 'PROBABILIDAD
          TABRIESGOS(POSI,4) = RS(4) 'IMPACTO

          FOR II=1 TO TTABRANRIESGO

              IF RS(2) =  TABRANRIESGOS(II,0) THEN
                 TABRIESGOS(POSI,2)    = TABRANRIESGOS(II,2) 'DESCRIPCION
                 TABRIESGOS(POSI,5)    = TABRANRIESGOS(II,3) 'COLOR
                 TABRIESGOS(POSI,6)    = II                  'RANGO
                 TABRANRIESGOS(ii,4)   = 1                   'CANTIDAD
                 EXIT FOR
              END IF

          NEXT

          TTABRIESGO = POSI

        RS.MOVENEXT
     LOOP



  END IF





  Strsql = "Select * From ORNIVRIESGO ORDER BY VALORDESDE "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOSRES(POSI,0) = RS(0) 'CODIGO

          TABRIESGOSRES(POSI,2) = CLNG(CDBL(RS(2)) * 100) 'VALOR DESDE
          TABRIESGOSRES(POSI,3) = CLNG(CDBL(RS(3)) * 100) 'VALOR HASTA

          FOR II=1 TO TTABRANRIESGO

              IF RS(1) =  TABRANRIESGOS(II,0) THEN
                 TABRIESGOSRES(POSI,1)    = TABRANRIESGOS(II,2) 'DESCRIPCION
                 TABRIESGOSRES(POSI,5)    = TABRANRIESGOS(II,3) 'COLOR
                 TABRIESGOSRES(POSI,6)    = II                  'RANGO
                 TABRANRIESGOS(ii,4)      = 1                   'CANTIDAD
                 EXIT FOR
              END IF

          NEXT

          TTABRIESGORES = POSI

        RS.MOVENEXT
     LOOP



  END IF


  Strsql = "     select d.secuen,COALESCE(e.descri,D.DESCRIACC), COALESCE(G.descri,G2.DESCRI), f.probabilidad,  "
  Strsql = Strsql & " f.impacto, f.riesgobruto, f.riesgoresidual,"
  Strsql = Strsql & " f.REQUIEREPLAN, f.INCPLAN, F.ACTIVIDAD, F.CODRIESGO, f.riesgoresidualNV, A.*, W.ESTADO, W.OBJID AS CODRSKPER, d.objid as CODRSKACTID,"
  strsql = strsql & " COALESCE(f.vulnerabilidad,f.vulnerabilidad) as vulnerabilidad,b.codvinobj"
  Strsql = Strsql & "     from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & "     inner join orrskcircuito w on (a.cobj = w.codobjeto AND W.ESTADO <> 4)"
  Strsql = Strsql & "     inner join objlnk x on (a.cobj = x.codobjeto and (x.ltipo = 'LRGVN' or x.ltipo = 'LRGCR'))"

  Strsql = Strsql & "     inner join objproced c on ( c.codobjeto = x.codobjetol and c.codproced = x.iden4 )"

  Strsql = Strsql & "     inner join objaccion d on ( C.codproced = d.codproced)  "

 if SFILSEL  = "SEC" then
  Strsql = Strsql & "     left join objsector h on ( C.codsector = h.objid)  "
 END IF

  Strsql = Strsql & "     LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt and x.ltipo = 'LRGVN' )"
  Strsql = Strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad and w.periodo = f.periodo)"
  Strsql = Strsql & "     left join orimpriesgo g on (f.itemriesgo = g.objid AND (b.codvinobj = 'VN' OR b.codvinobj = 'CR' ))"
  Strsql = Strsql & "     left join ORAMENAZAS g2 on (f.itemriesgo = g2.objid AND (b.codvinobj = 'AI'))"


    WKSPSELACC = ""

 IF SPROCMOD <> "WKF" THEN

    WKSPSELACC = " d.critico = 'SI' and  "

 END IF


  Strsql = Strsql & " where " & WKSPSELACC & " w.periodo = 0" & Speriodo & "   "

 IF SPROCESO <> "ALL" THEN

  Strsql = Strsql & " and ("
  Strsql = Strsql & " A.MC1 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC2 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC3 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC4 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC5 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC6 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC7 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.COBJ = " & SPROCESO & "  "
  Strsql = Strsql & " ) "

 END IF


  IF SFILSEL <> "" THEN

     IF SsubFILTRO = "" or SsubFILTRO = "ALL" THEN

        SsubFILTRO = "1"

     END IF



  if SFILSEL  = "PRO" then

     Strsql = Strsql & " AND f.probabilidad = " & TABPROBABILIDAD(SsubFILTRO,0) & "  "



  elseif SFILSEL  = "IMP" then

     Strsql = Strsql & " AND f.impacto = " & TABIMPACTO(SsubFILTRO,0) & "  "



  elseif SFILSEL  = "RBR" then

         Strsql = Strsql & " AND ("

         wspcon = ""

              FOR II=1 TO TTABRIESGO

               '''''   RESPONSE.WRITE " SF:" &  SsubFILTRO & "-TB:" & TABRIESGOS(II,6) & "-<BR>"

                  IF CLNG(TABRIESGOS(II,6)) =  CLNG(SsubFILTRO) THEN

                      SsubFILTROD = TABRIESGOS(II,3)
                      SsubFILTROH = TABRIESGOS(II,4)
                      Strsql = Strsql & wspcon & " ( f.probabilidad = 0" & SsubFILTROD & "  AND f.impacto = 0" & SsubFILTROH & ")  "

                      wspcon = " OR "

                  END IF

              NEXT



        Strsql = Strsql & " ) "


  elseif SFILSEL  = "RRE" then


              FOR II=1 TO TTABRIESGORES

               '''''   RESPONSE.WRITE " SF:" &  SsubFILTRO & "-TB:" & TABRIESGOS(II,6) & "-<BR>"

                  IF CLNG(TABRIESGOSRES(II,6)) =  CLNG(SsubFILTRO) THEN

                      SsubFILTROD = TABRIESGOSRES(II,2)
                      SsubFILTROH = TABRIESGOSRES(II,3)

                  END IF

              NEXT

   '''''''  Strsql = Strsql & " AND (f.riesgoresidual > '" & SsubFILTROD & "'  AND f.riesgoresidual < '" & SsubFILTROH & "')  "

             Strsql = Strsql & " AND ("

             Strsql  = Strsql & " NOT (f.riesgoresidual > 0" & replace(CDBL(SsubFILTROH / 100),",",".") & " ) AND "
             Strsql  = Strsql & " NOT (f.riesgoresidual < 0" & replace(CDBL(SsubFILTROD / 100),",",".") & " ) "

             Strsql = Strsql & " ) "


  elseif SFILSEL  = "RNV" then


   '''''  Strsql = Strsql & " AND f.riesgoresidualNV > '" & SsubFILTROD & "'  AND f.riesgoresidualNV < '" & SsubFILTROH & "'  "

              FOR II=1 TO TTABRIESGORES

               '''''   RESPONSE.WRITE " SF:" &  SsubFILTRO & "-TB:" & TABRIESGOS(II,6) & "-<BR>"

                  IF CLNG(TABRIESGOSRES(II,6)) =  CLNG(SsubFILTRO) THEN

                      SsubFILTROD = TABRIESGOSRES(II,2)
                      SsubFILTROH = TABRIESGOSRES(II,3)

                  END IF

              NEXT

             Strsql = Strsql & " AND ("

             Strsql  = Strsql & " NOT (f.riesgoresidualNV > 0" & replace(CDBL(SsubFILTROH / 100),",",".") & " ) AND "
             Strsql  = Strsql & " NOT (f.riesgoresidualNV < 0" & replace(CDBL(SsubFILTROD / 100),",",".") & " ) "

             Strsql = Strsql & " ) "


  elseif SFILSEL  = "SEC" then

     Strsql = Strsql & " AND H.CODSECTOR = " & SsubFILTRO & "  "

  elseif SFILSEL  = "CAT" then

     Strsql = Strsql & " AND G.AGRUPA = '" & SsubFILTRO & "'  "

     if SsubFILTRO2 <> "ALL"  then

        Strsql = Strsql & " AND G.CATEGO = '" & SsubFILTRO2 & "'  "

        if SsubFILTRO3 <> "ALL"  then

           Strsql = Strsql & " AND G.OBJID = '" & SsubFILTRO3 & "'  "

        END IF


     END IF


  end if





  END IF



 Strsql=Strsql & " ORDER BY  A.TSEC,D.SECUEN"
' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

           Celda = ""

 IF SPROCMOD = "OBJ" THEN
           Celda = Celda & ofv.gencelda("","VTG colspan=7 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY ACTIVOS EVALUADOS A<BR>LA BUSQUEDA SOLICITADA","si")
 ELSE
           Celda = Celda & ofv.gencelda("","VTG colspan=7 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY ACTIVIDADES EVALUADAS A<BR>LA BUSQUEDA SOLICITADA","si")
 END IF

           Celda = Celda & ofv.gencelda("","VT  style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"center","","","","","PERIODO:","si")
           Celda = Celda & ofv.gencelda("","VT colspan=3 style='background-color:#f7eedd;color:#888888' " & wkstylo1,"center","","","","",Speriododsc ,"si")

           Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


  ELSE
  posI = 0

  wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;' "
  wkstylo2 =" style='border-style:none;border-width:1;border-color:#aaaaaa;' "


        WKIPROCc = 0
        MAINPROC = 0
        MAINACTI = ""

           Celda = ""

 IF SPROCMOD = "OBJ" THEN
           Celda = Celda & ofv.gencelda("","VTI colspan=7 style='background-color:#AAAAAA;color:#ffffff;cursor:pointer;' " & wkstylo1 & "onclick=Open_All_Row();","","","","","","<span id='General'>(+)</span> MATRIZ DE IMPACTO DE AMENAZAS A ACTIVOS INFORMATICOS","si")
 ELSE
           Celda = Celda & ofv.gencelda("","VTI colspan=7 style='background-color:#AAAAAA;color:#ffffff;cursor:pointer;' " & wkstylo1 & "onclick=Open_All_Row();","","","","","","<span id='General'>(+)</span> MATRIZ DE IMPACTO DE RIESGO","si")
 END IF

           Celda = Celda & ofv.gencelda("","VT  style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"center","","","","","PERIODO:","si")
           Celda = Celda & ofv.gencelda("","VT colspan=3 style='background-color:#f7eedd;color:#888888' " & wkstylo1,"center","","","","",Speriododsc ,"si")


           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

  Do While (Not rs.EOF)

        WKIACTIV = ""

        WKIPROCD = ""
        WKIPROCN = 0
        WKIPROCC = RS(27) & ""

        FOR WKI=12 TO 19

            WKIPROC = RS(WKI) & ""

            IF WKIPROC <> "" THEN
               WKIPROCD = WKIPROCD & RS(WKI) & " / "
            END IF


        NEXT

        IF LEN(WKIPROCD) > 4 THEN
           WKIPROCD = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
        END IF


        WKIACTIV = WKIPROCC & RS(0) & RS(1) & ""

        IF MAINPROC <> WKIPROCC THEN                     ' CAMBIA PROCESO

         IF MAINPROC <> 0 THEN

           Celda = ""
           Celda = Celda & ofv.gencelda("","vt","","1%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","15%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","40%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

         END IF


           Celda = ""
           Celda = Celda & ofv.gencelda("","vt bgcolor='#ffffff' ","","1%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","UTI style='background-color:#8888ff;color:#ffffff' " & wkstylo1,"","","","","","PROCESO","si")
           Celda = Celda & ofv.gencelda("","uti COLSPAN=9 " & wkstylo1,"","","","","",WKIPROCD,"si")
           Fila  = FILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","30","","",Celda,"si")

           Celda = ""

           Celda = Celda & ofv.gencelda("","vt bgcolor='#ffffff' ","","1%","","","","&nbsp;","si")
 IF RS("codvinOBJ") = "AI" THEN
 'IF SPROCMOD = "OBJ" THEN

           Celda = Celda & ofv.gencelda("","uti colspan=2 " & wkstylo1,"","40%","","","","ACTIVIDAD","si")
           Celda = Celda & ofv.gencelda("","uti " & wkstylo1,"","40%","","","","AMENAZA ASOCIADA","si")
 ELSE
           Celda = Celda & ofv.gencelda("","uti colspan=2 " & wkstylo1,"","40%","","","","ACTIVIDAD","si")
           Celda = Celda & ofv.gencelda("","uti " & wkstylo1,"","40%","","","","RIESGO ASOCIADO","si")
 END IF


           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","PROBABILIDAD","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","IMPACTO","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","R.INHERENTE","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","R.RESIDUAL","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","REQ.PLAN","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","PLANIFICADO","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","RES.NUEVO","si")
           Fila  = FILA & ofv.genrow("","aaa bgcolor='#ddddff'  ","","","25","","",Celda,"si")






'''#F7DD88

           MAINPROC = WKIPROCC

        end if


    Celda = ""


      ' estado observacion

      Celdaobs  = ofv.gencelda("","vt","","1%","","","","&nbsp;","si")


      if RS("ESTADO") = 5 and SPROCMOD = "WKF" THEN


               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & RS("CODRSKPER")
      	       sAccion = sAccion & "&Tipobjeto=RSK&ID=" & RS("CODRSKACTID")  & "&ID2=" & RS("CODRSKPER")

      	       xsffun  = " onclick=""document.getElementById('DXA').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para la actividad.' ","","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")


      end if


      Celda = Celda & Celdaobs



                wspDATP     = 0
                wspDATI     = 0
                wspCOLORI   = "#EEEEEE"
                wspCOLORP   = "#EEEEEE"

              wspDATPXd = "---"
              wspDATIXd = "---"


             FOR II=1 TO TTABPROB
                 if TABPROBABILIDAD(ii,0) = RS(3) then
                    wspDATPXd   = TABPROBABILIDAD(ii,2) & ""
                    wspDATPX    = "0" & TABPROBABILIDAD(ii,1)
                    wspDATP     = CINT(wspDATPX)
                    wspCOLORP   = TABPROBABILIDAD(ii,5)
                    exit for
                 end if
             NEXT

              FOR II=1 TO TTABIMPA
                  if TABIMPACTO(ii,0) = RS(4) then
                     wspDATIXd   = TABIMPACTO(ii,2) & ""
                     wspDATIX    = "0" & TABIMPACTO(ii,1)
                     wspDATI     = CINT(wspDATIX)
                     wspCOLORI   = TABIMPACTO(ii,5)
                     exit for
                  end if
              NEXT



                wspDATrb    = RS(5)
                wspDATrr    = RS(6)
                wspDATrN    = RS(11)



'*****************************************************************************************************
'Comienzo de la entrada de la configuración de la tabla de cada uno de las filas de cada riesgo.     *
'*****************************************************************************************************

'' onclick=hide_row('fi" & posi &"_" & pos &"');


    sInput = "<span id='particular_" & posi & "'>(+)</span> " & RS(0) & "." & RS(1)
    Celda = Celda & ofv.gencelda("","UTi style='cursor:pointer;' colspan=2 " & wkstylo1 & "onclick=Open_Particular_Row('" &  posi & "');","","","","","",sInput,"si")

    'sInput = "(+) " & posi & " - " & RS(2) & "&nbsp;"
	sInput = RS(2) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")

    sInput = wspDATPXd
    Celda = Celda & ofv.gencelda("", wkstylof & " BGCOLOR='" & wspCOLORP & "' ","center","","","","",sInput ,"si")

    sInput = wspDATIXd
    Celda = Celda & ofv.gencelda("", wkstylof & " BGCOLOR='" & wspCOLORI & "' ","center","","","","",sInput,"si")

              wspDATrbXd = "---"
              wspCOLOR   = "#EEEEEE"
              wspDATrbvl = "N"

              FOR II=1 TO TTABRIESGO

                  IF TABRIESGOS(II,3) =  RS(3) AND TABRIESGOS(II,4) =  RS(4) THEN
                     wspDATrbXd   = TABRANRIESGOS(TABRIESGOS(II,6),2) & ""
                     wspCOLOR     = TABRANRIESGOS(TABRIESGOS(II,6),3)
                     wspDATrbvl   = "S"
                     EXIT FOR
                  END IF

              NEXT


    sInput = wspDATrbXd
    Celda = Celda & ofv.gencelda("", wkstylof & " BGCOLOR='" & wspCOLOR & "' ","center","","","","",sInput,"si")

              wspDATrrXd = "---"
              wspCOLOR   = "#EEEEEE"

           if wspDATrbvl = "S" then

              FOR II=1 TO TTABRIESGORES

                  IF CLNG(CDBL("0" & RS(6)) * 100) < TABRIESGOSRES(II,2)   OR CLNG(CDBL("0" & RS(6)) * 100) >  TABRIESGOSRES(II,3)  THEN

              wspDATrrXd = "---"
              wspCOLOR   = "#EEEEEE"

                  ELSE
                     wspDATrrXd = TABRANRIESGOS(TABRIESGOSRES(II,6),2) & ""
                     wspCOLOR   = TABRANRIESGOS(TABRIESGOSRES(II,6),3)
                     EXIT FOR
                  END IF

              NEXT

           end if

    sInput = wspDATrrXd
    Celda = Celda & ofv.gencelda("", wkstylof & " BGCOLOR='" & wspCOLOR & "' ","center","","","","",sInput,"si")

             wspcolor0A  = "---"

        if wspDATrbvl = "S" then


          if (RS(7) & "") = "SI"  then
             wspcolor0A  = sImageSI
          else
             wspcolor0A  = sImageNO
          end if

       end if




    sInput = wspcolor0A
    ' req plan
    Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput ,"si")

             wspcolor0A  = "---"

       if wspDATrbvl = "S" then

          if (RS(8) & "") = "SI"  then
             strsql= "select * from ORAGENDA  where  actividad='" & rs("actividad")  & "' and activriesgo='" & RS("CODRIESGO") &"' AND PERIODO=0" & SPERIODO
			    '  response.write strsql
			  set rsAGENDA = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			  sACTIVRIESGO=0
			  if not rs.eof  then sACTIVRIESGO=rsAGENDA(0)
              call ofv.cerrarconsulta(rsAGENDA)

             slnk="oragendaedit.asp?vision=PORMATIMPACTO&ACTIVRIESGO=" &  sACTIVRIESGO
             wspcolor0A  = "<a href=# onclick=window.open('"&slnk &"','myNewWin',',scrollbars=1,resizable=1,left=0,top=0,status=0,toolbar=0,width=900,height=700');  >" & sImageSI & "<a>"

          else
             wspcolor0A  = sImageNO
          end if

      end if




    sInput = wspcolor0A
    'planificado
    Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput ,"si")

              wspDATrnXd = "---"
              wspCOLOR   = "#EEEEEE"

           if wspDATrbvl = "S" then

              FOR II=1 TO TTABRIESGORES

                  IF CLNG(CDBL("0" & RS(11)) * 100) < TABRIESGOSRES(II,2)   OR CLNG(CDBL("0" & RS(11)) * 100) >  TABRIESGOSRES(II,3)  THEN

              wspDATrnXd = "---"
              wspCOLOR   = "#EEEEEE"

                  ELSE
                     wspDATrnXd = TABRANRIESGOS(TABRIESGOSRES(II,6),2) & ""
                     wspCOLOR   = TABRANRIESGOS(TABRIESGOSRES(II,6),3)
                     EXIT FOR
                  END IF

              NEXT

          end if


    sInput = wspDATrnXd
    Celda = Celda & ofv.gencelda("", wkstylof & " BGCOLOR='" & wspCOLOR & "' ","center","","","","",sInput,"si")



    Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")
    sInput = "&nbsp;"
    Celda =  ofv.gencelda("","UTi " & wkstylo2 &" colspan=3","","","","","",sInput,"si")
    Celda=celda & ofv.gencelda("", wkstylof & " BGCOLOR=#eeeeee colspan=8","left","","","","",replace(rs("vulnerabilidad")&"" ,chr(13),"<br>" ),"si")
    Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")




       if xTrColor = "#FFFFFF" then
          xTrColor = "#F9F9F9"
       ELSE
          xTrColor = "#FFFFFF"
       END IF


  if SFILSEL  = "CRT" OR SFILSEL  = "PRO" OR SFILSEL  = "IMP" OR SFILSEL  = "RBR" OR SFILSEL  = "RRE"  OR SFILSEL  = "RNV" then

  IF RS("codvinOBJ") = "AI" THEN

    sInput = "&nbsp;&nbsp;&nbsp;&nbsp;SIN SALVAGUARDAS MITIGANTES"
  else
   sInput = "&nbsp;&nbsp;&nbsp;&nbsp;SIN CONTROLES MITIGANTES"
  end if


    Strsql = "Select A.*	 From oractivrcontroles A "
    strsql = strsql & "  where a.periodo = 0" & Speriodo & " "
    strsql = strsql & " AND a.actividad = '" & RS(9) & "' AND A.CODRIESGO =  '" & RS(10) & "' "
    strsql = strsql & " ORDER BY A.tipo,A.SECUEN"

 ' response.write strsql & "<br>"
    set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RSX1.EOF THEN

       xTrColor2 = "#F7EEDD"

        FilaX1  = ""

        Celda  = ""
        IF RS("codvinOBJ") = "AI" THEN
          Celda  = Celda & ofv.GenCelda("","uti COLSPAN=2 style='background-color:#ddddff;'  "  & wkstylo1,"","","","","","SALVAGUARDAS","si")
		  Celda  = Celda & ofv.GenCelda("","UtI  style='background-color:#ddddff;'  "  & wkstylo1,"","","","","","CALIFICACION","si")
        ELSE
         Celda  = Celda & ofv.GenCelda("","uti COLSPAN=2 style='background-color:#ddddff;'  "  & wkstylo1,"","","","","","CONTROL MITIGANTE","si")
         Celda  = Celda & ofv.GenCelda("","UtI  style='background-color:#ddddff;'  "  & wkstylo1,"","","","","","CALIFICACION","si")
        END IF


        FilaX1  = FILAX1 & ofv.genrow("","","","","","","",Celda,"si")
     pos=0
     DO WHILE NOT RSX1.EOF

        Celda  = ""

          IF RSX1(10) = 1 THEN

             WKSPBGCOLOR = " BGCOLOR='#990000' "

          ELSE

             WKSPBGCOLOR = ""

          END IF

          Celda  = Celda & ofv.GenCelda("","UT " & WKSPBGCOLOR ,"","1%","","","","&nbsp;","si")
          IF RS ("codvinOBJ")<> "AI" THEN
           Strsql = "Select DESCRI From  ORCONTROLRIESGO WHERE OBJID = 0" & RSX1(7)
          ELSE
          Strsql = "Select DESCRI From ORsalvaguardas WHERE OBJID = 0" & RSX1(7)

          END IF

         ' response.write strsql & chr(13)
          set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          IF NOT RSX2.EOF THEN
             sInput2 = RSX2(0)

          ELSE
             sInput2 = "SIN ESPECIFICAR"
          END IF
'***********************************************************************************
'Configuración de la apertura de cada uno de los comentarios.                      *
'***********************************************************************************

          if len(trim(rsx1("descricont")&"")) > 0 then
           Celda  = Celda & ofv.GenCelda("","UTI style='cursor:pointer;' onclick=hide_row('fi" & posi &"_" & pos &"');","","69%","","","", "<span id='interna_fi_" & posi &"_" & pos &"'>(+)</span>" & sInput2,"si")
         '  Celda  = Celda & ofv.GenCelda("","UTI style='cursor:pointer;' onclick=hide_row('fi" & pos &"');","","69%","","","", "(+) " & sInput2,"si")
          else
           Celda  = Celda & ofv.GenCelda("","UTI ","","69%","","","",sInput2,"si")
          end if

          WSPCALIF = CLNG("0" &  RSX1(8))

          IF  WSPCALIF = 0 THEN


              Strsql = "Select DESCRI From ORNIVEFICIENCIA WHERE VALOR = 1"
              set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              IF NOT RSX2.EOF THEN
                 sInput2 = RSX2(0)
                ' sInput2=""
              ELSE
                 sInput2 = "SIN ESPECIFICAR"
              END IF

          ELSE

              Strsql = "Select DESCRI From ORNIVEFICIENCIA WHERE OBJID = 0" & RSX1(8)
              set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              IF NOT RSX2.EOF THEN
                 sInput2 = RSX2(0)

              ELSE
                 sInput2 = "SIN ESPECIFICAR"
              END IF

          END IF
          Celda  = Celda & ofv.GenCelda("","UTI","","30%","","","",sInput2,"si")


        FilaX1  = FILAX1 & ofv.genrow("","AAA bgcolor=" & xTrColor2 & " ","","","","","",Celda,"si")
        Celda  =  ofv.GenCelda("","UT "  ,"","1%","","","","&nbsp;","si") & ofv.GenCelda("","UTI colspan=2","","","","","",replace(rsx1("descricont")&"",chr(13),"<br>"),"si")

        if len(trim(rsx1("descricont")&"")) > 0 then
           FilaX1  = FILAX1 & ofv.genrow("","sss style='display:none;' id="&chr(34)&"fi" & posi &"_"& pos & chr(34),"","","","","",Celda,"si")
        end if

       ' FilaX1  = FILAX1 & ofv.genrow("","sss style='display:none;' id="&chr(34)&"fi" & pos & chr(34),"","","","","",Celda,"si")

       if xTrColor2 = "#F0F0F0" then
          xTrColor2 = "#F7EEDD"
       ELSE
          xTrColor2 = "#F0F0F0"
       END IF

     '    end if


     RSX1.MOVENEXT
     pos=pos+1


'***********************************************************************************
' Final de configuración de cada una de las celdas internaas.                      *
'***********************************************************************************

     LOOP


     sInput = ofv.GenTabla("","AAA " & wkstylo1 ,"","100% ","","0","","0","-1","","",FilaX1,"si")



    END IF

    Celda = ofv.gencelda("", "UT COLSPAN=3 ","","","","","","&nbsp;","si")

    Celda = Celda & ofv.gencelda("","UT COLSPAN=8 BGCOLOR='#F0F0F0' "  & wkstylo1 ,"","","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

   ' sInput = "&nbsp;"
   ' Celda = ofv.gencelda("","UT COLSPAN=11 " ,"","","","","",sInput,"si")

   ' Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

    FILA=FILA & muestraIndicadores(RS(10), posi)


  END IF




    posi = posi + 1
    rs.Movenext
  Loop

           Celda = ""
           Celda = Celda & ofv.gencelda("","vt","","1%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","15%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","40%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
          Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")


  END IF
  sHTML = sHTML & ofv.GenTabla("","aaa","","95% ","","0","","0","-1","","",Fila,"si")

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBSRSK()

'****************************************************************************************************************
'Comienzo de funciones de javascript para mostrar u ocultar los elementos que están ocultos
'****************************************************************************************************************
  sjavascrip=  chr(13) & "<script language=javascript>" & chr(13)
	  sjavascrip =sjavascrip & "function hide_row(nrow) { " & chr(13)
	  sjavascrip =sjavascrip & "	var v = document.getElementById(nrow);" & chr(13)
	  sjavascrip =sjavascrip & "	if (v.style.display=='none')" & chr(13)
	  sjavascrip =sjavascrip & "	{" & chr(13)
	  sjavascrip =sjavascrip & "		v.style.display ='block';"  & chr(13)
	  sjavascrip =sjavascrip & "	}"  & chr(13)
	  sjavascrip =sjavascrip & "	else"  & chr(13)
	  sjavascrip =sjavascrip & "	{" & chr(13)
	  sjavascrip =sjavascrip & "		v.style.display ='none';" & chr(13)
	  sjavascrip =sjavascrip & "	}" & chr(13)
	  sjavascrip=sjavascrip & "}"  & chr(13)
  sjavascrip=sjavascrip & "</script>"  & chr(13)  & chr(13)

  sjavascrip= sjavascrip & "<script language=javascript>"  & chr(13)
  	  sjavascrip=sjavascrip & "function Open_Particular_Row(nrow1)" & chr(13)
	  sjavascrip=sjavascrip & "{" & chr(13)
	'****************************************************************************
	'cambio para verificar si existen fi
	'****************************************************************************
	  sjavascrip=sjavascrip & "		var i_fi = 0;" & chr(13)
	  sjavascrip=sjavascrip & "		var fi = 'fi' + nrow1 + '_' + i_fi;" & chr(13)
	  sjavascrip=sjavascrip & "		if(document.getElementById(fi)==null)" & chr(13)
	  sjavascrip=sjavascrip & "		{" &  chr(13)
	  sjavascrip=sjavascrip & "			i_fi += 1;" &  chr(13)
	  sjavascrip=sjavascrip & "			var fi = 'fi' + nrow1 + '_' + i_fi;" & chr(13)
	  sjavascrip=sjavascrip & "			if(document.getElementById(fi)==null)" & chr(13)
	  sjavascrip=sjavascrip & "			{" & chr(13)
	  sjavascrip=sjavascrip & "				i_fi += 1;" &  chr(13)
	  sjavascrip=sjavascrip & "				var fi = 'fi' + nrow1 + '_' + i_fi;" & chr(13)
	  sjavascrip=sjavascrip & "				if(document.getElementById(fi)==null)" & chr(13)
	  sjavascrip=sjavascrip & "				{" & chr(13)
	  sjavascrip=sjavascrip & "				}" & chr(13)
	  sjavascrip=sjavascrip & "			}" & chr(13)
	  sjavascrip=sjavascrip & "		}" & chr(13)
	  
	  sjavascrip=sjavascrip & "		while(fi != '')" & chr(13)
	  sjavascrip=sjavascrip & "		{" & chr(13)
	  sjavascrip=sjavascrip & "			if(document.getElementById(fi)!=null)" & chr(13)
	  sjavascrip=sjavascrip & "			{" & chr(13)
	  sjavascrip=sjavascrip & "				var estado = document.getElementById('particular_' + nrow1).innerHTML;" & chr(13)
	  sjavascrip=sjavascrip & "				if(estado == '(+)')" & chr(13)
	  sjavascrip=sjavascrip & "				{" & chr(13)
	  sjavascrip=sjavascrip & "					document.getElementById(nombre).innerHTML = '(-)';" & chr(13)
	  sjavascrip=sjavascrip & "				}" & chr(13)
	  sjavascrip=sjavascrip & "				else" & chr(13)
	  sjavascrip=sjavascrip & "				{" & chr(13)
	  sjavascrip=sjavascrip & "					document.getElementById(nombre).innerHTML = '(+)';" & chr(13)
	  sjavascrip=sjavascrip & "				}" & chr(13)
	  sjavascrip=sjavascrip & "				hide_row(fi);" & chr(13)
	  sjavascrip=sjavascrip & "				i_fi += 1;" & chr(13)
	  sjavascrip=sjavascrip & "				fi = 'fi' + nrow1 + '_' + i_fi;" & chr(13)
	  sjavascrip=sjavascrip & "			}" & chr(13)
  	  sjavascrip=sjavascrip & "			else" & chr(13)
	  sjavascrip=sjavascrip & "			{" & chr(13)
	  sjavascrip=sjavascrip & "				fi='';" & chr(13)
	  sjavascrip=sjavascrip & "			}" & chr(13)
	  sjavascrip=sjavascrip & "		}" & chr(13)
	'****************************************************************************
	'cambio para verificar si existen xind
	'****************************************************************************
	  sjavascrip=sjavascrip & "		var i_xind = 0;" & chr(13)
	  sjavascrip=sjavascrip & "		var xind = 'xind' + nrow1 + '_' + i_xind;" & chr(13)
	  sjavascrip=sjavascrip & "		if(document.getElementById(xind)==null)" & chr(13)
	  sjavascrip=sjavascrip & "		{" & chr(13)
	  sjavascrip=sjavascrip & "			i_xind = 1;" & chr(13)
	  sjavascrip=sjavascrip & "			xind = 'xind' + nrow1 + '_' + i_xind;" & chr(13)
	  sjavascrip=sjavascrip & "			if(document.getElementById(xind)==null)" & chr(13)
	  sjavascrip=sjavascrip & "			{" & chr(13)
	  sjavascrip=sjavascrip & "				i_xind = 2;" & chr(13)
	  sjavascrip=sjavascrip & "				xind = 'xind' + nrow1 + '_' + i_xind;" & chr(13)
	  sjavascrip=sjavascrip & "				if(document.getElementById(xind)==null)" & chr(13)
	  sjavascrip=sjavascrip & "				{" & chr(13)
	  sjavascrip=sjavascrip & "					xind= '';" & chr(13)
	  sjavascrip=sjavascrip & "				}" & chr(13)
	  sjavascrip=sjavascrip & "			}" & chr(13)
	  sjavascrip=sjavascrip & "		}" & chr(13)
	  sjavascrip=sjavascrip & "		while(xind != '')" & chr(13)
	  sjavascrip=sjavascrip & "		{" & chr(13)
	  sjavascrip=sjavascrip & "			if(document.getElementById(xind)!=null)" & chr(13)
	  sjavascrip=sjavascrip & "			{" & chr(13)
	  sjavascrip=sjavascrip & "				hide_row(xind);" & chr(13)
	  sjavascrip=sjavascrip & "				i_xind += 1;" & chr(13)
	  sjavascrip=sjavascrip & "				xind = 'xind' + nrow1 + '_' + i_xind;" & chr(13)
	  sjavascrip=sjavascrip & "			}" & chr(13)
	  sjavascrip=sjavascrip & "			else" & chr(13)
	  sjavascrip=sjavascrip & "			{" & chr(13)
	  sjavascrip=sjavascrip & "				xind='';" & chr(13)
	  sjavascrip=sjavascrip & "			}" & chr(13)
	  sjavascrip=sjavascrip & "		}" & chr(13)
	  sjavascrip=sjavascrip & "}" & chr(13)
    sjavascrip=sjavascrip & "</script>" & chr(13) & chr(13)
	
  sjavascrip= sjavascrip & "<script language=javascript>"  & chr(13)
  	  sjavascrip=sjavascrip & "function Open_All_Row()" & chr(13)
	  sjavascrip=sjavascrip & "{" & chr(13)
	  sjavascrip=sjavascrip & "		var estado = document.all('General').innerText;" & chr(13)
	  sjavascrip=sjavascrip & "		if (estado == '(+)') " & chr(13) 
	  sjavascrip=sjavascrip & "		{" & chr(13)
	  sjavascrip=sjavascrip & "			document.all('General').innerText = '(-)';" & chr(13)
	  sjavascrip=sjavascrip & "		}" & chr(13)
	  sjavascrip=sjavascrip & "		else" & chr(13)
	  sjavascrip=sjavascrip & "		{" & chr(13)
	  sjavascrip=sjavascrip & "			document.all('General').innerText = '(+)';" & chr(13)
	  sjavascrip=sjavascrip & "		}" & chr(13)
	  sjavascrip=sjavascrip & "		var i = " & posi & ";" & chr(13)
	  'sjavascrip=sjavascrip & "		alert('hola');" & chr(13)
	  'sjavascrip=sjavascrip & "		var fi = 'fi' + i + '_0'" & chr(13)
	  sjavascrip=sjavascrip & "		for(j = 0; j < i; j++)" & chr(13)
	  sjavascrip=sjavascrip & "		{" & chr(13)
	  sjavascrip=sjavascrip & "			Open_Particular_Row(j, estado);" & chr(13)
	  'sjavascrip=sjavascrip & "			alert(j + ' , ' + i);" & chr(13)
	  sjavascrip=sjavascrip & "		}" & chr(13)
	  sjavascrip=sjavascrip & "}" & chr(13)
  sjavascrip=sjavascrip & "</script>" & chr(13) & chr(13)


'****************************************************************************************************************
'Fin de funciones de javascript para mostrar u ocultar los elementos que están ocultos
'****************************************************************************************************************
  sHTML = sHTML & "</body>"&sjavascrip &"</html>"


  function MuestraIndicadores(codriesgo, posi)
   DIM FILA
   DIM CELDA
   DIM RSXI

   Strsql = "Select B.DESCRIPCION,B.DESINDICADOR  From  ORRELAINDIRIESGO A,ORINDICADORES B where A.CODRIESGO='"& codriesgo &"'"
   STRSQL= STRSQL & " AND B.OBJID=A.idindicador"
  ' RESPONSE.WRITE STRSQL
   set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  IF NOT rsX1.EOF THEN
	  xTrColor2 = "#F7EEDD"

	   FilaX1  = ""

	   Celda  = ""
	   Celda  = Celda & ofv.GenCelda("","uti  style='background-color:#ddddff;'  "  & wkstylo1,"","","","","","NOMBRE DEL INDICADOR","si")
	 '  Celda  = Celda & ofv.GenCelda("","UtI  style='background-color:#ddddff;'  "  & wkstylo1,"","","","","","DESCRIPCION","si")

	   FilaX1  = FILAX1 & ofv.genrow("","","","","","","",Celda,"si")

	DO WHILE NOT rsX1.EOF

	   Celda  = ""


		 'Celda  = Celda & ofv.GenCelda("","UT " & WKSPBGCOLOR ,"","1%","","","","&nbsp;","si")

		 sInput2=rsX1(0)

		 if len(trim(rsX1(1)&"")) > 0 then
          Celda  = Celda & ofv.GenCelda("","UTI style='cursor:pointer;' onclick=hide_row('xind" & posi & "_" &  xposind  &"');","","69%","","","","<span id='interna_xind_" & posi &"_" & pos &"'>(+)</span> " & sInput2,"si")
         else
          Celda  = Celda & ofv.GenCelda("","UTI ","","69%","","","",sInput2,"si")
         end if


		' Celda  = Celda & ofv.GenCelda("","UTI ALIGN=LEFT ","","30%","","","",sInput2,"si")
		 FilaX1  = FILAX1 & ofv.genrow("","AAA bgcolor=" & xTrColor2 & " ","","","","","",Celda,"si")

		 sInput2=replace(rsX1(1)&"",chr(13),"<br>")
		 'IF LEN(sInput2&"") > 30  THEN sInput2=LEFT(sInput2,30) & "..."
		 Celda  =  ofv.GenCelda("","UTI","","70%","","","",sInput2,"si")

	     FilaX1  = FILAX1 & ofv.genrow("","AAA bgcolor=" & xTrColor2 & " style='display:none;' id=" & chr(34) & "xind" & posi & "_" & xposind & chr(34),"","","","","",Celda,"si")


	  rsX1.MOVENEXT
	  xposind=xposind+1
	LOOP

	CALL OFV.CERRARCONSULTA(rsX1)
	sInput = ofv.GenTabla("","AAA  " & wkstylo1 ,"","100% ","","0","","0","-1","","",FilaX1,"si")


   Celda = ofv.gencelda("", "UT COLSPAN=3 ","","","","","","&nbsp;","si")

   Celda = Celda & ofv.gencelda("","UT COLSPAN=8 BGCOLOR='#F0F0F0' "  & wkstylo1 ,"","","","","",sInput,"si")

   Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

   sInput = "&nbsp;"
   Celda = ofv.gencelda("","UT COLSPAN=11 " ,"","","","","",sInput,"si")

   Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")
   MuestraIndicadores=FILA

  ELSE


      sInput="&nbsp;&nbsp;&nbsp;&nbsp;SIN INDICADORES"
      Celda = ofv.gencelda("", "UT COLSPAN=3 ","","","","","","&nbsp;","si")

      Celda = Celda & ofv.gencelda("","UT COLSPAN=8 BGCOLOR='#F0F0F0' "  & wkstylo1 ,"","","","","",sInput,"si")

      Fila =  ofv.genrow("","","","","25","","",Celda,"si")

      sInput = "&nbsp;"
      Celda = ofv.gencelda("","UT COLSPAN=11 " ,"","","","","",sInput,"si")

      Fila = fila &  ofv.genrow("","","","","25","","",Celda,"si")
     MuestraIndicadores= FILA

  END IF
  end function
end if

Response.Write sHTML

%>