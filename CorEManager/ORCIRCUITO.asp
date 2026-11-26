<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 response.expires = 0

 sACTIVIDAD = request.querystring("ACTIVIDAD")
 scodmaster = request.querystring("codmaster") & ""
 stipmaster = request.querystring("tipmaster") & ""
 sOPCION    = request.querystring("OPCION")
 scodriesgo    = request.querystring("codriesgo")

 session("ModMain") = "ED"

''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)


  FUNCION = ""

  sHTML = ofv.MENUheader("",funcion)
  sHTML = sHTML & ofv.CorFormHeader



  Fila    = ""
  Celda   = ""
  sHref   = ""



'''''''''''''''''''''''''''''''''''''''''''



   wkproccod = "emdetaproc.asp?Nombre=" & request.querystring("Nombre")
   wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & scodmaster & "&tipmaster=" & stipmaster
   wkproccod = wkproccod & "&CODNIVEL=" & sCODNIVEL
   wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO

''   Session("ASPSTR") = wkproccod


''''''''''''''''''''''''''''''''''''''''''''
''' ESPECIFICACION DE PERIODOS PARA OR  ''''
''''''''''''''''''''''''''''''''''''''''''''

  wktabperiodo = ""
  Celda        = ""




   Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","10%","","","","PERIODO:","si")
   Celda = Celda & ofv.GenCelda("","uti ","","65%","","","",session("ORPERIODODESC"),"si")
   Celda = Celda & ofv.GenCelda("","ut  ","","20%","","","","&nbsp;","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   wktabperiodo = ofv.GenTabla("","","","100% ","100%","0","","-1","1","","",Fila,"si")




''''''''''''''''''''''''''''''''''''''''''''
 sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0233;</font>"

 wkproccod = "<a class=nt  href='" & Session("ASPSTR") & "' ><b>"
 wkproccod = wkproccod & sImage
 wkproccod = wkproccod & "</b></a>"

 AGREGAR = ""
 wkproccodN = ""


      Strsql = "Select A.* From OBJETOS A"
      Strsql = Strsql & " where  "
      Strsql = Strsql & "  A.CODOBJETO = '" & scodmaster & "' and A.TIPOBJETO = '" & stipmaster & "'  "


 '''  response.write strsql & "<br>" & chr(13)
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then


       WKPROCDESC = RS(3) & "&nbsp;"
       WKPROCVINC = CLNG("0" & RS(15))
       wspTIPO    = (rs(17) & "")


    END IF









  strsql = "select a.* from OBJACCION a "
  strsql = strsql & " WHERE a.CODACCION = '" &  SACTIVIDAD & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     sCODPROCED = RS(3)
     WKTIPOACC  = "M" & RS(29)



   IF wspTIPO = "XCR" OR wspTIPO = "VN" THEN


      'Acciones
      if rs(4) <> "" then
         Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & rs(4)
         set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rsAcciones.eof then
             sAccionM = rsAcciones(0)
             sModo    = rsAcciones(1)
         end if
         call ofv.cerrarconsulta(rsAcciones)
      end if

      WKSPTITLE = "Agregar riesgo asociado."

   ELSEIF wspTIPO = "CR"  THEN
      sAccionM = rs(28)
      WKSPTITLE = "Agregar riesgo asociado."
   else
      sAccionM = rs(28)

      WKSPTITLE = "Agregar amenaza asociada."

   end if





      WKPROCACCDESC = rs(10) & "&nbsp;.&nbsp;" & sAccionM


      sAccion = "ORCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
      sAccion = sAccion & "&Opcion=SI"


    IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4  AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 ) THEN

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='" & WKSPTITLE & "'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"

    END IF



  else

      WKPROCACCDESC = "no se encuentra " &  SACTIVIDAD


  END IF


           select case wspTIPO




                  case "VN"

                        sInput = "EVENTO VINCULADO EMANAGER"
                        sInput2 = "ACTIVIDAD"
                        wspTIPOCLC = "#008800"

                  case "CR"

                        sInput = "EVENTO MANUAL"
                        sInput2 = "ACTIVIDAD"
                        wspTIPOCLC = "#99CCDD"

                  case "AI"

                        sInput = "ACTIVO INFORMATICO"
                        sInput2 = "ACTIVO"
                        wspTIPOCLC = "#0000CC"

                  case else

                        sInput = "SIN DETERMINAR"
                        sInput2 = "----"
                        wspTIPOCLC = "#000000"

           end select


    sImagePRC = "<font face='corflaticon' size=2 color='" & wspTIPOCLC & "'  >&#060;</font>"

    WKDIVSECUEN = GENPROCSECUENEX()

    WKPROCDESC = WKDIVSECUEN & sInput & ":&nbsp;" & WKPROCDESC


    Celda = ""

   Celda = Celda & ofv.GenCelda("","Ut style='background-color:#f0f0f0;color:#990000;' ","CENTER","1%","","","",sImagePRC,"si")
   Celda = Celda & ofv.GenCelda("","ut style='background-color:#f0f0f0;color:#888888;' ","","85%","","","", WKPROCDESC,"si")
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","CENTER","1%","","","",wkproccodN,"si")
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","CENTER","1%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   TablaH = ofv.GenTabla("","aaa style='background-color:#f0eedd;border-style:solid;border-width:1;border-color:#000000;' valign='middle' align='center' ","","100% ","100%","0","","-1","1","","",Fila,"si")




   Celda = ""
   Celda = Celda & ofv.GenCelda("","uti style='Xbackground-color:#993333;color:#ffffff;' ","","","","","",sInput2 & ":&nbsp;" & WKPROCACCDESC,"si")
   Fila  = ofv.GenRow("","","","","30","","",Celda,"si")
   TablaH2 = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;background-color:#aaaaaa;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA RIESGOS ASOCIADOS - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN



        strsql = "Insert into ORACTIVRIESGO "
        strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
        strsql = strsql & " SECUEN, CODRIESGO, ITEMRIESGO, PROBABILIDAD, IMPACTO, "
        strsql = strsql & " RIESGOBRUTO, RIESGORESIDUAL, RIESGORESIDUALNV, VULNERABILIDAD,REQUIEREPLAN) "
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & session("ORPERIODO") & "','" & sACTIVIDAD & "','" & sCODPROCED & "',"
        strsql = strsql &  "'" & SCODMASTER & "','" & STIPMASTER & "',"
        strsql = strsql & "'10','" & GTkey(SCODMASTER) & "',"
        strsql = strsql & "'0','0','0','0','0','0',' ','NO'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)


 END IF


 IF SOPCION = "DEL" THEN

        strsql = "Delete From ORACTIVRIESGO Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "Delete From ORACTIVRCONTROLES Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "Delete From ORRELAINDIRIESGO Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)





 END IF

''''''''''''''''''''''''''''''''''''''''''''''''''

   Fila = ""

   wspcolor1  = "#990000"
   wspcolor2  = "#009900"
   wspcolor3  = "#D7CF77"
   wspcolor4  = "#EEEEEE"
   sImageDAP  = "<font face='corflaticon' size=3 color='#FFFFFF'>3</font>"

  DIM TABRANRIESGOS(10,5)


  Strsql = "Select * From ORRANGORIESGO ORDER BY VALOR DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRANRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRANRIESGOS(POSI,1) = RS(1) 'VALOR
          TABRANRIESGOS(POSI,2) = RS(2) 'DESCRIPCION
          TABRANRIESGOS(POSI,3) = RS(3) 'COLOR

          TTABRANRIESGO = POSI

        RS.MOVENEXT
     LOOP



  END IF



   'Strsql = "Select a.*,b.objid as Control,c.objid as indicador From ORACTIVRIESGO a"
    Strsql = "Select a.* From ORACTIVRIESGO a"
  ' strsql =strsql & " left  join  oractivrcontroles b on (b.periodo=a.periodo and b.actividad=a.ACTIVIDAD and b.codriesgo=a.CODRIESGO)"
 '  strsql=strsql & "  left join  ORRELAINDIRIESGO c  on  (c.CODRIESGO=a.CODRIESGO)"
   strsql= strsql & " where a.PERIODO = 0" & session("ORPERIODO") & " AND a.ACTIVIDAD = '" & SACTIVIDAD & "' "
   Strsql = Strsql & " ORDER BY a.SECUEN "

' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     Celda = ""

   IF wspTIPO = "CR" OR wspTIPO = "VN" THEN
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN RIESGOS ASOCIADOS A LA ACTIVIDAD","si")
   else
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN AMENAZAS ASOCIADAS AL ACTIVO","si")
   end if


     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

     Tabla = "<CENTER><BR><BR><BR><BR>" & Tabla & "</CENTER>"

  else


       wkcoddelete = "N"
       wktitdelete = " "
       sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
       sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"

    xTrColor = "#FFFFFF"

    CeldaK  = ""
    Celda  = ""

   IF wspTIPO = "CR" OR wspTIPO = "VN" THEN
      Celda  = Celda & ofv.GenCelda("","uti colspan=3  style='background-color:#888888;color:#ffffff;' ","","","","","","RIESGOS ASOCIADOS","si")
   else
      Celda  = Celda & ofv.GenCelda("","uti colspan=3  style='background-color:#888888;color:#ffffff;' ","","","","","","AMENAZAS ASOCIADAS" ,"si")
   end if

    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#888888;color:#ffffff;' ","CENTER","5%","","","","INHERENTE&nbsp;&nbsp;&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#888888;color:#ffffff;' ","CENTER","5%","","","","RESIDUAL","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#555555;color:#ffffff;' ","CENTER","5%","","","","PLAN&nbsp;&nbsp;&nbsp;&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#555555;color:#ffffff;' ","CENTER","5%","","","","RESIDUAL NUEVO","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#555555;color:#ffffff;' ","CENTER","5%","","","","INDICADOR","si")
    FilaN  = filaN & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

    do until rs.eof


       wspcolor  = "#336699"
       sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

       wkproccod = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD
       wkproccod = wkproccod & "&RIESGO=" & RS(7) & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster


       WKCODRIESGO = "NO"



   IF wspTIPO = "CR" OR wspTIPO = "VN" THEN

       sRIESGODESC = "<FONT COLOR='#880000' style='border-style:solid;border-width:1;border-COLOR:#993333;' TITLE='ESTABLECER RIESGO ASOCIADO.'  > ESTABLECER RIESGO ASOCIADO </FONT>"

      'CODIGO DE RIESGO
       Strsql = "Select * From ORIMPRIESGO Where OBJID = 0" & rs(8)
       set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  not rsAcciones.eof then
           sCODNEMO    = rsAcciones(1)
           sDESCRI     = rsAcciones(2)
           sRIESGODESC = "<U TITLE='ACCEDER / MODIFICAR RIESGO ASOCIADO.' >" & sCODNEMO & "&nbsp;-&nbsp;" & sDESCRI & "</U>"


       end if
       call ofv.cerrarconsulta(rsAcciones)

       wktitdelete = "Confirma la eliminacion del riesgo asociado???."

   else

       sRIESGODESC = "<FONT COLOR='#880000' style='border-style:solid;border-width:1;border-COLOR:#993333;' TITLE='ESTABLECER AMENAZA ASOCIADA.'  > ESTABLECER AMENAZA ASOCIADA </FONT>"

      'CODIGO DE RIESGO
       Strsql = "Select * From ORAMENAZAS Where OBJID = 0" & rs(8)
       set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  not rsAcciones.eof then
           sCODNEMO    = rsAcciones(1)
           sDESCRI     = rsAcciones(2)
           sRIESGODESC = "<U TITLE='ACCEDER / MODIFICAR AMENAZA ASOCIADA.' >" & sCODNEMO & "&nbsp;-&nbsp;" & sDESCRI & "</U>"


       end if
       call ofv.cerrarconsulta(rsAcciones)

       wktitdelete = "Confirma la eliminacion de la amenaza asociada???."


   end if




    IF RS(8) > 0 THEN WKCODRIESGO = "SI"


       sHref     = wkproccod
       sLink     = ofv.GenLink("","nt",sHref,"","","",sRIESGODESC,"si")

       Celdak = ""
       Celda  = ""


       wkcoddelete = "S"
       'response.write  isnull(rs("control")) & "--" &  isnull (rs("indicador"))
      ' if  not isnull(rs("control")) or not isnull (rs("indicador")) then wkcoddelete = "N"

       if wkcoddelete = "S" AND SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 ) then
          sImageEL = sImageELS
          sformobj = "FRMRGA" & rs(0)
       '''   sAccion    = "respproceso.asp?codobjeto=" & rs(7) & "&tipobjeto=" & "RGA" & ""

          sAccion = "ORCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
          sAccion = sAccion & "&Opcion=DEL&CODRIESGO=" & RS(7)

          formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
          sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "


       else
          wktitdelete = "No se puede eliminar"
          sImageEL = sImageELN
          formobj    = ""
          sfunobj    = ""
       end if




       Celda     = ""
       Celda = Celda & ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2%","","","",sImageEL,"si") & formobj


       Celda  = Celda & ofv.GenCelda("","ut","CENTER","1%","","","",sImageP,"si")
       Celda  = Celda & ofv.GenCelda("","UT","","65%","","","",sLink,"si")


       IF WKCODRIESGO = "NO" THEN

          wspcolor0A  = wspcolor4
          wspcolor0B  = wspcolor4
          wspcolor0P  = wspcolor4
          wspcolor0Bb  = wspcolor4

          wspDATA     = "---"
          wspDATB     = "---"
          wspDATP     = "---"
          wspDATBb     = "---"

             wspmontobruto    = "0"
             wspmontoresidual = "0"
             wspmontoresidaj  = "0"
             wspDATAdsc       = "-----"
             wspDATArrdsc     = "-----"
             wspDATArradsc    = "-----"

             wspmontobrn      = ""
             wspmontoren      = ""
             wspmontoran      = ""

       ELSE




             wspDATA     = rs(11)
             wspDATB     = rs(12)
             wspDATP     = rs(14)
             wspDATBB     = rs(17)

             wspDATP      = 0
             wspDATI      = 0

             wspmontobruto    = "0"
             wspmontoresidual = "0"
             wspmontoresidaj  = "0"

             wspmontobrn      = ""
             wspmontoren      = ""
             wspmontoran      = ""

             wspDATAdsc       = "-----"
             wspDATArrdsc     = "-----"
             wspDATArradsc    = "-----"


             wspporc = 100

             Strsql  = "Select VALOR,porcentaje From ORNIVPROBOCURRENCIA where OBJID = 0" & RS(9)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then
                wspDATPX    = "0" & RS1(0)
                wspDATP     = CINT(wspDATPX)
                wspporc     = CINT("0" & RS1(1))
             END IF

             wspmonto = 0

             Strsql  = "Select VALOR,impdesde,imphasta From ORNIVIMPACTO where OBJID = 0" & RS(10)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then
                wspDATIX    = "0" & RS1(0)
                wspDATI     = CINT(wspDATIX)
                wspmonto    = clng((Clng("0" & RS1(1)) + Clng("0" & RS1(2))) / 2)
             END IF


            if wspDATP > 0 and wspDATI > 0 then





             Strsql  = "Select descri From ORNIVRIESGOBRUTO where NIVPROB =  0" & RS(9) & " AND NIVIMPA = 0" & RS(10) & " "
 '''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
                       wspDATAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATAdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT


             END IF


             Strsql  = "Select descri From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "

           ''  Strsql  = "Select descri From ORNIVriesgo where valorhasta > 0" & replace(cstr(wspDATB & ""),",",".")
  '' response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
                       wspDATArrdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArrdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT

              ''  wspDATArrdsc    = RS1(0)
             END IF


             Strsql  = "Select descri From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "

           ''  Strsql  = "Select descri From ORNIVriesgo where valorhasta > 0" & replace(cstr(wspDATBB & ""),",",".")
 ''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
                       wspDATArrAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArrAdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT


           ''     wspDATArradsc    = RS1(0)
             END IF


             wspDATA     = cdbl(wspDATA)
             wspDATB     = cdbl(wspDATB)
             wspDATBB    = cdbl(wspDATBB)

             IF wspDATB = 0 THEN wspDATB = 1
             IF wspDATBB = 0 THEN wspDATBB = 1

             wspmontobruto    = clng(wspmonto * wspporc / 100)
             wspmontobrn      = "-----<br>$&nbsp;" & wspmontobruto


             wspDATM          = cdbl(round(wspDATA / wspDATB,2))
             wspDATMB         = cdbl(round(wspDATA / wspDATBB,2))


             wspmontoresidual = clng(wspmontobruto / wspDATM)
             wspmontoresidaj  = clng(wspmontobruto / wspDATMB)

       '''      wspmontoren      = "-----<br>$&nbsp;" & wspmontoresidual & "-" & wspDATM & "-" & wspDATB
       '''      wspmontoran      = "-----<br>$&nbsp;" & wspmontoresidaj  & "-" & wspDATMB & "-" & wspDATBB


             wspmontoren      = "-----<br>$&nbsp;" & wspmontoresidual
             wspmontoran      = "-----<br>$&nbsp;" & wspmontoresidaj




            end if






          if rs(14) = "SI"  then

             wspcolor0P  = wspcolor2

          else

             wspcolor0P  = wspcolor4

          end if




       END IF



       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#888888;COLOR:#000000;"
       wsstyloF   = wsstyloF & "background-color:#f0f0f0;' "

       Celdak  = Celdak & ofv.GenCelda("","Vt rowspan=2 " & wsstyloF,"CENTER","5%","","","",wspDATAdsc & "<br>" & wspmontobrn,"si")

       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#000077;COLOR:#888888;"
       wsstyloF   = wsstyloF & "background-color:#f0f0f0;' "

       Celdak  = Celdak & ofv.GenCelda("","Vt  rowspan=2  " & wsstyloF,"CENTER","5%","","","",wspDATArrdsc & "<br>" & wspmontoren,"si")

       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#888888;COLOR:#888888;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";' "

       Celdak  = Celdak & ofv.GenCelda("","Ut  rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")


       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#993333;COLOR:#888888;"
       wsstyloF   = wsstyloF & "background-color:#f0f0f0;' "

       Celdak  = Celdak & ofv.GenCelda("","Vt  rowspan=2  " & wsstyloF,"CENTER","5%","","","",wspDATArradsc & "<br>" & wspmontoran,"si")


        wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#888888;COLOR:#888888;"

       if wspDATP <> "---" then
       if  wspDATP > 0 and wspDATI  then
        strsql= "SELECT  * FROM ORRELAINDIRIESGO WHERE  CODRIESGO='" & rs(7) &"'"
		SET RSREL =ofv.crearconsultaEx(StrSql,cn,1,volver)

        IF not RSREL.EOF THEN
	     wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
	    else
	     wspcolor0P="#993333"
	     wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
	    end if
        call ofv.cerrarconsulta(RSREL)
        url="'ORASIGNAINDICADOR.asp?opcion=CONSULTA&riesgo=" &rs(7) & "' ,'mywindow','status=0,toolbar=0,width=900,height=400'"
        Celdak  = Celdak & ofv.GenCelda("","Ut  onclick="&chr(34)&"window.open("&url&");"&chr(34)&"  rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")
       else
         wspcolor0P="#993333"
	     wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
         celdak=celdak & ofv.GenCelda("","Ut  rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")
       end if
       else
        url="'ORASIGNAINDICADOR.asp?opcion=CONSULTA&riesgo=" &rs(7) & "' ,'mywindow','status=0,toolbar=0,width=900,height=400'"
        wspcolor0P="#993333"
	    wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
        celdak=celdak  & ofv.GenCelda("","Ut  onclick="&chr(34)&"window.open("&url&");"&chr(34)&"    rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")
       end if




       Celda = Celda & Celdak
       FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")

       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","uti colspan=2   ","","","","","","&nbsp;","si")
       Celda  = Celda & ofv.GenCelda("","uti   style='border-style:solid;border-width:1;border-COLOR:#bbbbbb;' ","","","","","",rs(13) & "&nbsp;","si")
       FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

      rs.movenext
    loop

    Tabla = "<BR>" & ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")

  end if

    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"


  sHTML = sHTML & "<SCRIPT language=" & Chr(34) & "VBScript" & Chr(34) & ">" & Chr(13)

  sHTML = sHTML & " function Abre() " & Chr(13)
  sHTML = sHTML & " if window.event.button=2  then  " & Chr(13)
'  sHTML = sHTML & "    if window.event.srcElement.tagName = ""A"" or window.event.srcElement.tagName = ""IMG""  or window.event.srcElement.tagName = ""FONT"" then  " & Chr(13)
'  sHTML = sHTML & "       window.alert(""dato "" & window.event.srcElement.ID & "" tag "" & window.event.srcElement.tagName ) " & Chr(13)
  sHTML = sHTML & "       Editar(window.event.srcElement.ID) " & Chr(13)
 ' sHTML = sHTML & "    end if " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " end function " & Chr(13) & Chr(13)

  sHTML = sHTML & " function Editar(Elemento) " & Chr(13)
  sHTML = sHTML & " if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
  sHTML = sHTML & "    exit function " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " dato = " & Chr(34) & "edicion2.asp?"



  sHTML = sHTML & "ID=" & scodmaster & "-" & STIPMASTER & "X-X" & Chr(34) & Chr(13)


  sHTML = sHTML & " window.showModalDialog "
  sHTML = sHTML & "dato,window," & Chr(34) & "dialogWidth:600px;dialogHeight:360px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)



  sHTML = sHTML & " end function " & Chr(13)

  sHTML = sHTML & "</SCRIPT> " & Chr(13)

'''  sHTML = sHTML & FormuP



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''''''''''''''''''''''''

  Celda = ""

  Celda = Celda & "<div id='DX01' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:110px;display:block;"
  Celda = Celda & "width:95%;height:30;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
  Celda = Celda & "background-color:#f9f9f9;"
  Celda = Celda & "' >"
  Celda = Celda & wktabperiodo
  Celda = Celda & "</div>"

  Celda = Celda & "<div id='DX0' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:140px;display:block;"
  Celda = Celda & "width:95%;height:30;overflow:auto;"
  Celda = Celda & "background-color:#f9f9f9;"
  Celda = Celda & "' >"
  Celda = Celda & TablaH
  Celda = Celda & "</div>"

  Celda = Celda & "<div id='DX2' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:170px;display:block;"
  Celda = Celda & "width:95%;height:70%;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
  Celda = Celda & "background-color:" & session("APLBGD") & ";"
  Celda = Celda & "' >"
  Celda = Celda & TablaH2
  Celda = Celda & TABLA
  Celda = Celda & "</div>"

  sHTML = sHTML & Celda


'''''''''''''''''''''''''''''''''''''''''''''



  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,"no","")

  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML

 FUNCTION GENPROCSECUENEX()

    GENPROCSECUENEX = ""

    Strsql = "Select * From CSRELPATHDESCRI  where COBJ = " & scodmaster
 '' response.write strsql & "<br>"
    set Wrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT Wrs.eof then

       WKITOPE = 8

       FOR WKI=0 TO 7

           WKIPROC = WRS(WKI) & ""

           IF WKIPROC <> "" THEN
              WKITOPE = WKI
              EXIT FOR
           END IF


       NEXT

       WKICCOL = 7 - WKITOPE

       WKINCOL = 0

       FilaW = ""

           Celda = ""
           Celda2 = ""

       FOR WKI=WKITOPE TO 7




           Celda = Celda & Celda2 & WRS(WKI)

           Celda2 = "<b><font color='#00aa00' size=2 >&#0124;</font></b>"



       NEXT


       IF WKITOPE < 8 THEN

          GENPROCSECUENEX = Celda & Celda2

       END IF



    END IF


 END FUNCTION


%>

