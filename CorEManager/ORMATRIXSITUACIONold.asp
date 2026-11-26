<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


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

  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

     if Speriodo = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL')  " 
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")

        end if

      else

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & Speriodo
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           Speriododsc              = RS(1)

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

  wkstylof = "AAA style='font-family:arial;font-size:6pt;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' "
  wkstylof2 = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "

''''b.codvinobj = 'VN' AND


 IF SPROCMOD = "OBJ" THEN


  Strsql = "     select B.INDOBJ,B.DESCRI, g.descri, f.probabilidad,  "
  Strsql = Strsql & " f.impacto, f.riesgobruto, f.riesgoresidual,"
  Strsql = Strsql & " f.REQUIEREPLAN, f.INCPLAN, F.ACTIVIDAD, F.CODRIESGO, f.riesgoresidualNV, A.* "
  Strsql = Strsql & "     from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"

  Strsql = Strsql & "     left join oractivriesgo f on (convert(nvarchar(100),b.codobjeto) = f.actividad ) "
  Strsql = Strsql & "     left join oramenazas g on (f.itemriesgo = g.objid) "

  Strsql = Strsql & " where f.periodo = 0" & Speriodo & " "

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



 ELSE 


  Strsql = "     select d.secuen,ISNULL(e.descri,D.DESCRIACC), g.descri, f.probabilidad,  "
  Strsql = Strsql & " f.impacto, f.riesgobruto, f.riesgoresidual,"
  Strsql = Strsql & " f.REQUIEREPLAN, f.INCPLAN, F.ACTIVIDAD, F.CODRIESGO, f.riesgoresidualNV, A.*, W.ESTADO, W.OBJID AS CODRSKPER, d.objid as CODRSKACTID "
  Strsql = Strsql & "     from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & "     inner join orrskcircuito w on (a.cobj = w.codobjeto AND W.ESTADO <> 4)"
  Strsql = Strsql & "     inner join objlnk x on (a.cobj = x.codobjeto and (x.ltipo = 'LRGVN' or x.ltipo = 'LRGCR'))"

  Strsql = Strsql & "     inner join objproced c on ( c.codobjeto = x.codobjetol and c.codproced = x.iden4 )"

''  Strsql = Strsql & "         (a.cobj = c.codobjeto) "
''  Strsql = Strsql & "           or ( c.codproced = ( select iden4 from objlnk where codobjeto = a.cobj) ) "
''  Strsql = Strsql & "                                ) "

  Strsql = Strsql & "     inner join objaccion d on ( C.codproced = d.codproced)  "

 if SFILSEL  = "SEC" then
  Strsql = Strsql & "     left join objsector h on ( C.codsector = h.objid)  "
 END IF

  Strsql = Strsql & "     LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"
  Strsql = Strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad and w.periodo = f.periodo)"
  Strsql = Strsql & "     left join orimpriesgo g on (f.itemriesgo = g.objid)"


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

     Strsql = Strsql & " AND f.probabilidad = " & SsubFILTRO & "  "

     

  elseif SFILSEL  = "IMP" then

     Strsql = Strsql & " AND f.impacto = " & SsubFILTRO & "  "



  elseif SFILSEL  = "RBR" then


          if SsubFILTRO = "1"  then
             SsubFILTROD = "-1"
             SsubFILTROH = "6"

          elseif SsubFILTRO = "2"  then
             SsubFILTROD = "5"
             SsubFILTROH = "11"
          elseif SsubFILTRO = "3" then
             SsubFILTROD = "10"
             SsubFILTROH = "16"
          elseif SsubFILTRO = "4"  then
             SsubFILTROD = "15"
             SsubFILTROH = "21"
          elseif SsubFILTRO = "5" then
             SsubFILTROD = "20"
             SsubFILTROH = "26"
          end if  

     Strsql = Strsql & " AND f.riesgobruto > '" & SsubFILTROD & "'  AND f.riesgobruto < '" & SsubFILTROH & "'  "



  elseif SFILSEL  = "RRE" then

          if SsubFILTRO = "1"  then
             SsubFILTROD = "-1"
             SsubFILTROH = "6"

          elseif SsubFILTRO = "2"  then
             SsubFILTROD = "5"
             SsubFILTROH = "11"
          elseif SsubFILTRO = "3" then
             SsubFILTROD = "10"
             SsubFILTROH = "16"
          elseif SsubFILTRO = "4"  then
             SsubFILTROD = "15"
             SsubFILTROH = "21"
          elseif SsubFILTRO = "5" then
             SsubFILTROD = "20"
             SsubFILTROH = "26"
          end if  

     Strsql = Strsql & " AND f.riesgoresidual > '" & SsubFILTROD & "'  AND f.riesgoresidual < '" & SsubFILTROH & "'  "


  elseif SFILSEL  = "RNV" then

          if SsubFILTRO = "1"  then
             SsubFILTROD = "-1"
             SsubFILTROH = "6"

          elseif SsubFILTRO = "2"  then
             SsubFILTROD = "5"
             SsubFILTROH = "11"
          elseif SsubFILTRO = "3" then
             SsubFILTROD = "10"
             SsubFILTROH = "16"
          elseif SsubFILTRO = "4"  then
             SsubFILTROD = "15"
             SsubFILTROH = "21"
          elseif SsubFILTRO = "5" then
             SsubFILTROD = "20"
             SsubFILTROH = "26"
          end if  

     Strsql = Strsql & " AND f.riesgoresidualNV > '" & SsubFILTROD & "'  AND f.riesgoresidualNV < '" & SsubFILTROH & "'  "


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


 END IF

 ''' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 

           Celda = ""

 IF SPROCMOD = "OBJ" THEN
           Celda = Celda & ofv.gencelda("","VTG colspan=5 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY ACTIVOS ASOCIADOS A<BR>LA BUSQUEDA SOLICITADA","si")
 ELSE
           Celda = Celda & ofv.gencelda("","VTG colspan=5 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY ACTIVIDADES ASOCIADAS A<BR>LA BUSQUEDA SOLICITADA","si")
 END IF

           Celda = Celda & ofv.gencelda("","VT  style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"center","","","","","PERIODO:","si")
           Celda = Celda & ofv.gencelda("","VT colspan=3 style='background-color:#f7eedd;color:#888888' " & wkstylo1,"center","","","","",Speriododsc ,"si")

           Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


  ELSE
  posI = 0

                wspDATP     = 0 
                wspDATI     = 0 
                wspDATrb    = 0  
                wspDATrr    = 0 
                wspDATrN    = 0 

                wspCANT     = 0 


                wspTPCANT     = 0 

                wspTPDATP     = 0 
                wspTPDATI     = 0 
                wspTPDATrb    = 0  
                wspTPDATrr    = 0 
                wspTPDATrN    = 0 

                wspTTCANT     = 0 

                wspTTDATP     = 0 
                wspTTDATI     = 0 
                wspTTDATrb    = 0  
                wspTTDATrr    = 0 
                wspTTDATrN    = 0 

  wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;' "


        WKIPROCc = 0
        MAINPROC = 0
        MAINACTI = ""

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTI colspan=5 style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"","","","","","SITUACION DEL PROCESO","si")
           Celda = Celda & ofv.gencelda("","VT  style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"center","","","","","PERIODO:","si")
           Celda = Celda & ofv.gencelda("","VT colspan=3 style='background-color:#f7eedd;color:#888888' " & wkstylo1,"center","","","","",Speriododsc ,"si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

           Celda = ""
           Celda = Celda & ofv.gencelda("","UTI COLSPAN=4 ROWSPAN=2 style='background-color:#8888ff;color:#ffffff' " & wkstylo1,"","","","","","PROCESO","si")
           Celda = Celda & ofv.gencelda("","uti COLSPAN=5 style='background-color:#8888ff;color:#ffffff' " & wkstylo1,"","","","","","INDICADORES","si")
           Fila  = FILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","30","","",Celda,"si")

           Celda = ""

           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","PROBABILIDAD","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","IMPACTO","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","R.BRUTO","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","R.RESIDUAL","si")
           Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"","10%","","","","RESID.AJ.","si")

           Fila  = FILA & ofv.genrow("","aaa bgcolor='#ddddff'  ","","","25","","",Celda,"si")


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
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")









           Celda = ""
           Celda = Celda & ofv.gencelda("","uti COLSPAN=4 " & wkstylo1,"","","","","",WKIPROCD,"si")


          if wspTPCANT > 0 then


                wspTPDATP     = wspTPDATP   / wspTPCANT
                wspTPDATI     = wspTPDATI   / wspTPCANT
                wspTPDATrb    = wspTPDATrb  / wspTPCANT
                wspTPDATrr    = wspTPDATrr  / wspTPCANT
                wspTPDATrN    = wspTPDATrN  / wspTPCANT


          if wspTPDATP = 0  then

             wspcolor0A  = sImagevn

          elseif wspTPDATP = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTPDATP = 2  then

             wspcolor0A  = sImagevb

          elseif wspTPDATP = 3  then

             wspcolor0A  = sImagevm

          elseif wspTPDATP = 4  then

             wspcolor0A  = sImageva

          elseif wspTPDATP = 5  then

             wspcolor0A  = sImagevMa

          end if  



           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")



          if wspTPDATI = 0  then

             wspcolor0A  = sImagevn

          elseif wspTPDATI = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTPDATI = 2  then

             wspcolor0A  = sImagevb

          elseif wspTPDATI = 3  then

             wspcolor0A  = sImagevm

          elseif wspTPDATI = 4  then

             wspcolor0A  = sImageva

          elseif wspTPDATI = 5  then

             wspcolor0A  = sImagevMa

          end if  

           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")

          if wspTPDATrb = 0  then
             wspcolor0A  = sImagevn
          elseif wspTPDATrb < 6  then
             wspcolor0A  = sImagevNUL
          elseif wspTPDATrb < 11  then
             wspcolor0A  = sImagevb
          elseif wspTPDATrb < 16  then
             wspcolor0A  = sImagevm
          elseif wspTPDATrb < 21  then
             wspcolor0A  = sImageva
          elseif wspTPDATrb > 20  then
             wspcolor0A  = sImagevMa
          end if  


          if wspTPDATrb = 0  then

             wspcolor0A  = sImagevn

          elseif wspTPDATrb = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTPDATrb = 2  then

             wspcolor0A  = sImagevb

          elseif wspTPDATrb = 3  then

             wspcolor0A  = sImagevm

          elseif wspTPDATrb = 4  then

             wspcolor0A  = sImageva

          elseif wspTPDATrb = 5  then

             wspcolor0A  = sImagevMa

          end if  
 

           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")



          if wspTPDATrr = 0  then
             wspcolor0A  = sImagevn
          elseif wspTPDATrr < 6  then
             wspcolor0A  = sImagevNUL
          elseif wspTPDATrr < 11  then
             wspcolor0A  = sImagevb
          elseif wspTPDATrr < 16  then
             wspcolor0A  = sImagevm
          elseif wspTPDATrr < 21  then
             wspcolor0A  = sImageva
          elseif wspTPDATrr > 20  then
             wspcolor0A  = sImagevMa
          end if  


          if wspTPDATrr = 0  then

             wspcolor0A  = sImagevn

          elseif wspTPDATrr = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTPDATrr = 2  then

             wspcolor0A  = sImagevb

          elseif wspTPDATrr = 3  then

             wspcolor0A  = sImagevm

          elseif wspTPDATrr = 4  then

             wspcolor0A  = sImageva

          elseif wspTPDATrr = 5  then

             wspcolor0A  = sImagevMa

          end if  

           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")


          if wspTPDATrN = 0  then
             wspcolor0A  = sImagevn
          elseif wspTPDATrN < 6  then
             wspcolor0A  = sImagevNUL
          elseif wspTPDATrN < 11  then
             wspcolor0A  = sImagevb
          elseif wspTPDATrN < 16  then
             wspcolor0A  = sImagevm
          elseif wspTPDATrN < 21  then
             wspcolor0A  = sImageva
          elseif wspTPDATrN > 20  then
             wspcolor0A  = sImagevMa
          end if  

          if wspTPDATrN = 0  then

             wspcolor0A  = sImagevn

          elseif wspTPDATrN = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTPDATrN = 2  then

             wspcolor0A  = sImagevb

          elseif wspTPDATrN = 3  then

             wspcolor0A  = sImagevm

          elseif wspTPDATrN = 4  then

             wspcolor0A  = sImageva

          elseif wspTPDATrN = 5  then

             wspcolor0A  = sImagevMa

          end if  


           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")


         else

           sInput = "------------"
           Celda = Celda & ofv.gencelda("", wkstylof & " colspan=5 ","center","","","","",sInput,"si")

         end if

           Fila  = FILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","30","","",Celda,"si")

                wspTPCANT     = 0 

                wspTPDATP     = 0 
                wspTPDATI     = 0 
                wspTPDATrb    = 0  
                wspTPDATrr    = 0 
                wspTPDATrN    = 0 

          else


         END IF



'''#F7DD88

           MAINPROC = WKIPROCC

        end if


    Celda = ""


      ' estado observacion
      
      Celdaobs  = ofv.gencelda("","vt","","1%","","","","&nbsp;","si")

        


      Celda = Celda & Celdaobs



                wspDATP     = 0 
                wspDATI     = 0 
                wspDATrb    = 0  
                wspDATrr    = 0 
                wspDATrN    = 0 


             Strsql  = "Select VALOR,descri From ORNIVPROBOCURRENCIA where OBJID = 0" & RS(3)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then
                wspDATPXd   = RS1(1) & "" 
                wspDATPX    = "0" & RS1(0)    
                wspDATP     = CINT(wspDATPX)    
             END IF

             Strsql  = "Select VALOR,descri From ORNIVIMPACTO where OBJID = 0" & RS(4)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then
                wspDATIXd   = RS1(1) & ""   
                wspDATIX    = "0" & RS1(0)    
                wspDATI     = CINT(wspDATIX)    
             END IF

    if wspDATI = 0 or wspDATP = 0 then

    else 

                wspDATrb    = RS(5)    
                wspDATrr    = RS(6)    
                wspDATrN    = RS(11)    



                wspTPCANT   = wspTPCANT + 1
                wspTTCANT   = wspTTCANT + 1







    sInput = RS(0) & "." & RS(1)
    Celda = Celda & ofv.gencelda("","UTi colspan=2 " & wkstylo1,"","","","","",sInput,"si")

    sInput = RS(2) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")


          if wspDATP = 0  then

             wspcolor0A  = sImagevn

          elseif wspDATP = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspDATP = 2  then

             wspcolor0A  = sImagevb

          elseif wspDATP = 3  then

             wspcolor0A  = sImagevm

          elseif wspDATP = 4  then

             wspcolor0A  = sImageva

          elseif wspDATP = 5  then

             wspcolor0A  = sImagevMa

          end if  



    sInput = wspcolor0A
    Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")



          if wspDATI = 0  then

             wspcolor0A  = sImagevn

          elseif wspDATI = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspDATI = 2  then

             wspcolor0A  = sImagevb

          elseif wspDATI = 3  then

             wspcolor0A  = sImagevm

          elseif wspDATI = 4  then

             wspcolor0A  = sImageva

          elseif wspDATI = 5  then

             wspcolor0A  = sImagevMa

          end if  

    sInput = wspcolor0A
    Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")

          if wspDATrb = 0  then
             wspcolor0A  = sImagevn
             wspDATrb    = 0    
          elseif wspDATrb < 6  then
             wspcolor0A  = sImagevNUL
             wspDATrb    = 1    
          elseif wspDATrb < 11  then
             wspcolor0A  = sImagevb
             wspDATrb    = 2    
          elseif wspDATrb < 16  then
             wspcolor0A  = sImagevm
             wspDATrb    = 3    
          elseif wspDATrb < 21  then
             wspcolor0A  = sImageva
             wspDATrb    = 4    
          elseif wspDATrb > 20  then
             wspcolor0A  = sImagevMa
             wspDATrb    = 5    
          end if  



    sInput = wspcolor0A
    Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")



          if wspDATrr = 0  then
             wspcolor0A  = sImagevn
             wspDATrr    = 0    
          elseif wspDATrr < 6  then
             wspcolor0A  = sImagevNUL
             wspDATrr    = 1    
          elseif wspDATrr < 11  then
             wspcolor0A  = sImagevb
             wspDATrr    = 2    
          elseif wspDATrr < 16  then
             wspcolor0A  = sImagevm
             wspDATrr    = 3    
          elseif wspDATrr < 21  then
             wspcolor0A  = sImageva
             wspDATrr    = 4    
          elseif wspDATrr > 20  then
             wspcolor0A  = sImagevMa
             wspDATrr    = 5    
          end if  




    sInput = wspcolor0A
    Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")


          if wspDATrN = 0  then
             wspcolor0A  = sImagevn
             wspDATrn    = 0    
          elseif wspDATrN < 6  then
             wspcolor0A  = sImagevNUL
             wspDATrn    = 1    
          elseif wspDATrN < 11  then
             wspcolor0A  = sImagevb
             wspDATrn    = 2    
          elseif wspDATrN < 16  then
             wspcolor0A  = sImagevm
             wspDATrn    = 3    
          elseif wspDATrN < 21  then
             wspcolor0A  = sImageva
             wspDATrn    = 4    
          elseif wspDATrN > 20  then
             wspcolor0A  = sImagevMa
             wspDATrn    = 5    
          end if  




    sInput = wspcolor0A
    Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")






    Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")



       if xTrColor = "#FFFFFF" then
          xTrColor = "#F9F9F9" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF




                wspTPDATP     = wspTPDATP   + wspDATP
                wspTPDATI     = wspTPDATI   + wspDATI
                wspTPDATrb    = wspTPDATrb  + wspDATrb
                wspTPDATrr    = wspTPDATrr  + wspDATrr
                wspTPDATrN    = wspTPDATrN  + wspDATrN


                wspTTDATP     = wspTTDATP   + wspDATP
                wspTTDATI     = wspTTDATI   + wspDATI
                wspTTDATrb    = wspTTDATrb  + wspDATrb
                wspTTDATrr    = wspTTDATrr  + wspDATrr
                wspTTDATrN    = wspTTDATrN  + wspDATrN

    end if



 

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
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")


                wspTTDATP     = round(wspTTDATP   / wspTTCANT,0)
                wspTTDATI     = round(wspTTDATI   / wspTTCANT,0)
                wspTTDATrb    = round(wspTTDATrb  / wspTTCANT,0)
                wspTTDATrr    = round(wspTTDATrr  / wspTTCANT,0)
                wspTTDATrN    = round(wspTTDATrN  / wspTTCANT,0)


           Celda = ""
           Celda = Celda & ofv.gencelda("","uti COLSPAN=4 " & wkstylo1,"","","","","","SITUACION DEL PROCESO","si")


          if wspTTDATP = 0  then

             wspcolor0A  = sImagevn

          elseif wspTTDATP = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTTDATP = 2  then

             wspcolor0A  = sImagevb

          elseif wspTTDATP = 3  then

             wspcolor0A  = sImagevm

          elseif wspTTDATP = 4  then

             wspcolor0A  = sImageva

          elseif wspTTDATP = 5  then

             wspcolor0A  = sImagevMa

          end if  



           sInput = wspcolor0A & wspTTDATP
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")



          if wspTTDATI = 0  then

             wspcolor0A  = sImagevn

          elseif wspTTDATI = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTTDATI = 2  then

             wspcolor0A  = sImagevb

          elseif wspTTDATI = 3  then

             wspcolor0A  = sImagevm

          elseif wspTTDATI = 4  then

             wspcolor0A  = sImageva

          elseif wspTTDATI = 5  then

             wspcolor0A  = sImagevMa

          end if  

           sInput = wspcolor0A & wspTTDATI
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")




         if wspTTDATrb = 0  then

             wspcolor0A  = sImagevn

          elseif wspTTDATrb = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTTDATrb = 2  then

             wspcolor0A  = sImagevb

          elseif wspTTDATrb = 3  then

             wspcolor0A  = sImagevm

          elseif wspTTDATrb = 4  then

             wspcolor0A  = sImageva

          elseif wspTTDATrb = 5  then

             wspcolor0A  = sImagevMa

          end if  
 

           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")



          if wspTTDATrr = 0  then

             wspcolor0A  = sImagevn

          elseif wspTTDATrr = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTTDATrr = 2  then

             wspcolor0A  = sImagevb

          elseif wspTTDATrr = 3  then

             wspcolor0A  = sImagevm

          elseif wspTTDATrr = 4  then

             wspcolor0A  = sImageva

          elseif wspTTDATrr = 5  then

             wspcolor0A  = sImagevMa

          end if  

           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")



          if wspTTDATrN = 0  then

             wspcolor0A  = sImagevn

          elseif wspTTDATrN = 1  then

             wspcolor0A  = sImagevNUL

          elseif wspTTDATrN = 2  then

             wspcolor0A  = sImagevb

          elseif wspTTDATrN = 3  then

             wspcolor0A  = sImagevm

          elseif wspTTDATrN = 4  then

             wspcolor0A  = sImageva

          elseif wspTTDATrN = 5  then

             wspcolor0A  = sImagevMa

          end if  


           sInput = wspcolor0A
           Celda = Celda & ofv.gencelda("", wkstylof,"center","","","","",sInput,"si")





           Fila  = FILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","30","","",Celda,"si")


  END IF
  sHTML = sHTML & ofv.GenTabla("","aaa ","","95% ","","0","","0","-1","","",Fila,"si")

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""
  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBSRSK()

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>