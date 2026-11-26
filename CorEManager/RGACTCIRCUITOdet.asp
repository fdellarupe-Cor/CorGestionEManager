<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 response.expires = 0
' response.write request.querystring
 sACTIVIDAD = request.querystring("ACTIVIDAD")
 sRIESGO    = request.querystring("RIESGO")
 scodmaster = request.querystring("codmaster") & ""
 stipmaster = request.querystring("tipmaster") & ""
 sOPCION    = request.querystring("OPCION")
 sCONTROLES = request.querystring("CONTROLES")
 SCODCONTROL = request.querystring("CODCONTROL")
 SCONTROLTIPO = request.querystring("CONTROLTIPO")
 volver2 = ofv.vueltaasp2("")

 IF SCONTROLTIPO = "" THEN SCONTROLTIPO = 0

 session("ModMain") = "ED"

''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)


  FUNCION = ""

  sHTML = ofv.formheader("")
'  sHTML = sHTML & ofv.CorFormHeader



  Fila    = ""
  Celda   = ""
  sHref   = ""


 wkproccod = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=SI&CONTROLTIPO=" & SCONTROLTIPO
 wkproccod = wkproccod & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster

 VUELTA = "../" & REPLACE(wkproccod,"&","_")

 wkproccod = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&CONTROLES=SI&CONTROLTIPO=" & SCONTROLTIPO
 wkproccod = wkproccod & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster

 VUELTA2 = "../" & REPLACE(wkproccod,"&","_")


'''''''''''''''''''''''''''''''''''''''''''

 WKTIPOOBJ = "ACT"


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

 IF sCONTROLES = "SI" THEN

    wkproccod = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RIESGO=" & SRIESGO & "&RECALC=SI&CONTROLTIPO=" & SCONTROLTIPO
    wkproccod = wkproccod & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster

 ELSE

    wkproccod = "RGACTCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&CONTROLTIPO=" & SCONTROLTIPO
    wkproccod = wkproccod & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster


 END IF



 wkproccod = "<a class=nt  href='" & wkproccod & "' ><b>"
 wkproccod = wkproccod & sImage
 wkproccod = wkproccod & "</b></a>"

 AGREGAR = ""


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







'''''''''''''''''


  strsql = "SELECT A.* FROM RGACTIVOPRTY A "
  strsql = strsql & " WHERE A.CODOBJETO =  0" & SACTIVIDAD & "  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     sCODPROCED = RS(0)
     WKTIPOACC  = "M" & RS(0)

     sTIPOBJETO = RS(2)

     sACTVALUACION = RS(6)
     sACTFAMILY    = RS(7)

   IF wspTIPO = "XCR" OR wspTIPO = "VN" THEN

      WKSPTITLE = "Agregar riesgo asociado."

   ELSEIF wspTIPO = "CR"  THEN

      WKSPTITLE = "Agregar riesgo asociado."
   else

      WKSPTITLE = "Agregar amenaza asociada."

   end if


             Strsql = "Select * From RGACTFAMILY WHERE OBJID = " & sACTFAMILY
             set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rsX2.eof then 

                WKTDESCRI   = RSX2(1)

                WKTCONFIDEN = CINT("0" & RSX2(2))
                WKTINTEGRID = CINT("0" & RSX2(3))
                WKTDISPONIB = CINT("0" & RSX2(4))

''''''''' ANALIZA EL IMPACTO POR CADA DIMENSION

                WKTCONFIDENV = sACTVALUACION * WKTCONFIDEN / 100
                WKTINTEGRIDV = sACTVALUACION * WKTINTEGRID / 100
                WKTDISPONIBV = sACTVALUACION * WKTDISPONIB / 100






             END IF





      WKPROCACCDESC = "ACTIVO INFORMATICO"


      sAccion = "RGACTCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
      sAccion = sAccion & "&Opcion=SI" & "&CONTROLTIPO=" & SCONTROLTIPO



  else

      WKPROCACCDESC = "no se encuentra " &  SACTIVIDAD


  END IF
''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA CONTROLES MITIGANTES - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN



        strsql = "Insert into ORACTIVRCONTROLES "
        strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
        strsql = strsql & " SECUEN, CODRIESGO, ITEMCONTROL, CALIFMITIGANTE,TIPO,MODOCONT) "
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & session("ORPERIODO") & "','" & sACTIVIDAD & "','" & sCODPROCED & "','" & sACTIVIDAD & "','" & sTIPOBJETO & "',"
        strsql = strsql & "'10','" & SRIESGO & "',"
        strsql = strsql & "'0','0','" & SCONTROLTIPO & "','1'"
        strsql = strsql & ")"
   '     RESPONSE.WRITE STRSQL
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


 END IF

 IF SOPCION = "DEL" THEN


        strsql = "Delete From ORACTIVRCONTROLES Where OBJID = 0" & SCODCONTROL
        call ofv.crearconsultaEx(StrSql,cn,1,volver)




 END IF

''''''''''''''''''''''''''''''''''''''''''''''''''

   Fila = ""
   FilaN = ""
   wkproccodN = ""

   wspcolor1  = "#990000"
   wspcolor2  = "#009900"
   wspcolor3  = "#D7CF77"
   wspcolor4  = "#EEEEEE"

   wspcolorV1  = "#00dd00"
   wspcolorV2  = "#6699cc"
   wspcolorV3  = "#F7DD88"
   wspcolorV4  = "#f7aa22"
   wspcolorV5  = "#dd0000"

          sImageO1   = "&#0253;"
          sImageO2   = "&#0254;"

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



   Strsql = "Select A.*, b.actividad AS A1 ,c.actividad AS A2 From RGACTANALISIS A "
   strsql = strsql & " left join (select distinct actividad, CODRIESGO "
   strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 0) b "
   strsql = strsql & " on (a.actividad = b.actividad AND A.CODRIESGO = B.CODRIESGO )"
   strsql = strsql & " left join (select distinct actividad, CODRIESGO "
   strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 1) c "
   strsql = strsql & " on (a.actividad = c.actividad AND A.CODRIESGO = c.CODRIESGO )"
   strsql = strsql & " where A.CODRIESGO =  '" & sRIESGO & "' "

'''response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     if (rs("a1") & "") = "" then



                 strsql = "Insert into ORACTIVRCONTROLES "
  	         strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
  	         strsql = strsql & " SECUEN, CODRIESGO, ITEMCONTROL, CALIFMITIGANTE,TIPO,MODOCONT) "
  	         strsql = strsql & " Values ("
  	         strsql = strsql &  "'" & session("ORPERIODO") & "','" & sACTIVIDAD & "','" & sCODPROCED & "','" & sACTIVIDAD & "','" & sTIPOBJETO & "',"
  	         strsql = strsql & "'10','" & SRIESGO & "',"
  	         strsql = strsql & "'0','0','" & SCONTROLTIPO & "','1'"
  	         strsql = strsql & ")"
  	      '   RESPONSE.WRITE STRSQL
                 call ofv.crearconsultaEx(StrSql,cn,1,volver)


     end if

  end if

  call ofv.cerrarconsulta(rs)


  Strsql = "Select A.*, b.actividad AS A1 ,c.actividad AS A2 From RGACTANALISIS A "
     strsql = strsql & " left join (select distinct actividad, CODRIESGO "
     strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 0) b "
     strsql = strsql & " on (a.actividad = b.actividad AND A.CODRIESGO = B.CODRIESGO )"
     strsql = strsql & " left join (select distinct actividad, CODRIESGO "
     strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 1) c "
     strsql = strsql & " on (a.actividad = c.actividad AND A.CODRIESGO = c.CODRIESGO )"
     strsql = strsql & " where A.CODRIESGO =  '" & sRIESGO & "' "

'  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     Fila = ""
     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTE LA AMENAZA ASOCIADA AL ACTIVO","si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
     Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")
     Tabla = "<CENTER><BR><BR><BR><BR>" & Tabla & "</CENTER>"

  else



      sAccion = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&CONTROLES=SI"
      sAccion = sAccion & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster
      sAccion = sAccion & "&Opcion=SI" & "&CONTROLTIPO=" & SCONTROLTIPO


   IF SCONTROLES = "SI" AND SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 )  THEN

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='AGREGAR MITIGANTES.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"

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


 ''  Celda = ""
 ''  Celda = Celda & ofv.GenCelda("","uti style='Xbackground-color:#993333;color:#ffffff;' ","","","","","",sInput2 & ":&nbsp;" & WKPROCACCDESC,"si")
 ''  Fila  = ofv.GenRow("","","","","30","","",Celda,"si")
 ''  TablaH2 = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;background-color:#aaaaaa;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

   TablaH2 = ""

    xTrColor = "#FFFFFF"

    CeldaK  = ""
    Celda  = ""

  IF SCONTROLES = "SI" THEN

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''' POR CONTROLES MITIGANTES
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       Fila = ""

    if (SCONTROLTIPO & "") = "0" then

       wcnttipo = " AND a.TIPO = '" & SCONTROLTIPO & "' "


    else

       wcnttipo = ""


    end if

    Strsql = "Select A.* From oractivrcontroles A "
    strsql = strsql & "  where a.periodo = 0" & session("ORPERIODO") & " "
    strsql = strsql & " AND a.actividad = '" & SACTIVIDAD & "' AND A.CODRIESGO =  '" & sRIESGO & "' " & wcnttipo
    strsql = strsql & " ORDER BY a.TIPO, A.objid "

'response.write strsql & "<br>"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if rs.eof then


       Celda = ""


       Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN SALVAGUARDAS MITIGANTES PARA LA AMENAZA ASOCIADA","si")
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
       Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

       Tabla = "<CENTER><BR><BR><BR><BR>" & Tabla & "</CENTER>"

     else

       wkcoddelete = "N"
       wktitdelete = " "
       sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
       sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"

       sImageVIG    = "<font face='corflaticon' size=3 color='#00AA00'>&#097;</font>"
       sImageREM    = "<font face='corflaticon' size=3 color='#F7AA22'>&#0114;</font>"

       Celda = ""
       Celda  = Celda & ofv.GenCelda("","uti colspan=5  style='background-color:#336699;color:#ffffff;'  ","","","","","","CALIFICACION DE LA AMENAZA ASOCIADA","si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

       wktitdeleteS = "Confirma la eliminacion de la salvaguarda mitigante???."


       Celda = ""
       Celda  = Celda & ofv.GenCelda("","uti colspan=2  style='background-color:#336699;color:#ffffff;'  ","","","","","","SALVAGUARDAS MITIGANTES","si")
       Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#336699;color:#ffffff;'  ","CENTER","","","","","TIPO","si")
       Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#336699;color:#ffffff;'  ","CENTER","","","","","CALIFICACION","si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

'response.write strsql
       do until rs.eof

          sForm = "formPER" & RS(0)
          Fila   = ""
          Celda  = ""

          wkcoddelete = "S"
          wkmodif     = "N"

          if SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 ) and ( ((SCONTROLTIPO & "") = "0") or ((SCONTROLTIPO & "") = "1" and (rs(10) & "" = "1") ) ) then
             wkmodif = "S"
          end if

          if wkcoddelete = "S" AND wkmodif = "S" then

             wktitdelete = wktitdeleteS

             sImageEL = sImageELS
             sformobj = "FRMRGA" & rs(0)

             sAccion    = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&OPCION=DEL&CODCONTROL=" & RS(0) & "&CONTROLTIPO=" & SCONTROLTIPO
             sAccion    = sAccion & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster & "&CONTROLES=SI"

             formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
             sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "


          else
             wktitdelete = "No se puede eliminar. "
             sImageEL = sImageELN
             formobj    = ""
             sfunobj    = ""

             IF (SCONTROLTIPO & "") = "1" and (rs(10) & "" = "0") THEN

                IF (RS(9) & "") = "0" THEN

                   wktitdelete = "Control reemplazado o desafectado. "
                   sImageEL = sImagerem

                ELSE

                   wktitdelete = "Control vigente. "
                   sImageEL = sImagevig

                END IF


             END IF



           end if


         Celda     = ""
         Celda = Celda & ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " style='border-style:solid;border-width:1;border-color:#888888;' ","","1%","","","",sImageEL,"si")


         if wkmodif = "S" then


          '''  Strsql = "Select * From ORSALVAGUARDAS  WHERE GRUPO = 0" & sACTFAMILY & " ORDER BY DESCRI"

            Strsql = " SELECT A.*,B.ACTFAMILYCOD FROM ORSALVAGUARDAS A "
            Strsql = Strsql & " INNER JOIN RGACTFAMILYSVRSK B ON (A.OBJID = B.ACTRSKCOD) "
            Strsql = Strsql & " WHERE B.ACTFAMILYCOD = 0" & sACTFAMILY & " ORDER BY A.DESCRI"


            sInput = ofv.GenerarCombo(sForm,RS(7).NAME,RS(7).value,Strsql,cn,0,1,"")
         else
            sInput = "-----"

            strsql = "Select descri From ORSALVAGUARDAS where OBJID = 0" & RS(7)
            set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rsniv.eof then sInputx = rsniv(0)
            call ofv.cerrarconsulta(rsniv)

            sInput = sInputx
         end if


          Celda  = Celda & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","","70%","","","",sInput,"si")

          if wkmodif = "S" then
             'scomK = "1" & chr(9) & "2" & chr(9) & "3"
             'scomD = "AUTOMATICO" & chr(9) & "SEMIAUTOMATICO" & chr(9) & "MANUAL"
             'sInput = ofv.gencomboFX(sForm,rs(11).name,rs(11),scomK,scomD," ")
			
			Strsql = "Select * From RGTIPOSALVAGUARDA ORDER BY DESCRIPCION"
             sInput = ofv.GenerarCombo(sForm,RS(11).NAME,RS(11).value,Strsql,cn,0,1,"")
          else
             'sInput = "-----"

             'if (rs(11) & "") = "1" then
             '   sInput = "AUTOMATICO"
             'elseif (rs(11) & "") = "2" then
             '   sInput = "SEMIAUTOMATICO"
             'elseif (rs(11) & "") = "3" then
             '   sInput = "MANUAL"
             'end if
			 sInput = "-----"

             strsql = "Select DESCRIPCION From RGTIPOSALVAGUARDA where OBJID = 0" & RS(11)
             set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rsniv.eof then sInputx = rsniv(0)
             call ofv.cerrarconsulta(rsniv)

             sInput = sInputx
          end if

          Celda  = Celda & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

          if wkmodif = "S" then
             Strsql = "Select * From RGACTEFICIENCIASV ORDER BY VALORD"
             sInput = ofv.GenerarCombo(sForm,RS(8).NAME,RS(8).value,Strsql,cn,0,1,"")
          else
             sInput = "-----"

             strsql = "Select DESCRIPCION From RGACTEFICIENCIASV where OBJID = 0" & RS(8)
             set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rsniv.eof then sInputx = rsniv(0)
             call ofv.cerrarconsulta(rsniv)

             sInput = sInputx

          end if

          Celda  = Celda & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")
          Fila   = Fila & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")

          Celda  = ""
          Celda  = Celda & ofv.GenCelda("","uti colspan=4  style='background-color:#aaaaaa;color:#000000;'  ","","","","","","DESCRIPCION","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

          nCol = 80

           wdescri = " "

          if  not  isnull(rs(12)) then
              wdescri = rs(12) & ""
          end if

          t3b = "<font class=at><textarea class=dt "
          t3b = t3b & "name=" & rs(12).name & " cols=" & nCol & " rows=2 "
          if wkmodif = "S" then
              t3b = t3b & " onblur=" & sForm & ".submit() "
              t3b = t3b & " onkeypress=" & chr(34)
              t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
              t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
          else
              t3b = t3b & " readonly >"
          end if
          xdatos =  wdescri
          t3b    = t3b & xdatos & "</textarea></font>"
          SINPUT = t3b

          Celda  = ""
          Celda  = Celda & ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

          if (rs(10) & "") = "0" then
              sInput = "EXISTENTE"

              IF (SCONTROLTIPO & "") = "1" THEN

                Celda  = Celda & ofv.GenCelda("","uti  style='border-style:solid;border-width:1;border-color:#888888;COLOR:#000000;' ","","","","","",sInput,"si")

                scomK = "10" & chr(9) & "0"
                scomD = "VIGENTE" & chr(9) & "REEMPLAZADO"
                sInput = ofv.gencomboFX(sForm,rs(9).name,rs(9),scomK,scomD," ")

                Celda  = Celda & ofv.GenCelda("","ut  style='border-style:solid;border-width:1;border-color:#888888;COLOR:#000000;' ","","","","","",sInput,"si")

              ELSE

                Celda  = Celda & ofv.GenCelda("","uti COLSPAN=2 style='border-style:solid;border-width:1;border-color:#888888;COLOR:#000000;' ","","","","","",sInput,"si")
 
              END IF

          else
              sInput = "<B>NUEVO</B>"
              Celda  = Celda & ofv.GenCelda("","uti COLSPAN=2 style='border-style:solid;border-width:1;border-color:#888888;COLOR:#000000;' ","","","","","",sInput,"si")

          end if

          Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")


          Celda  = ""
          SINPUT = "&nbsp;"
          Celda  = Celda & ofv.GenCelda("","uti colspan=4 style='border-TOP-style:SOLID;border-TOP-width:2;border-TOP-color:#000000;' ","","","","","",sInput,"si")
          Fila = fila & ofv.GenRow("","AAA BGCOLOR='#ffffff' ","","","25","","",Celda,"si")


         IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 )  THEN
            sAccion = "asp/actualizar.asp?fname=" & rs(0)
            sAccion = sAccion & "&tabla=ORACTIVRCONTROLES&volver=" & vuelta2
         ELSE
            sAccion = ""
         END IF

          FilaN = filaN & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si") & formobj


          if xTrColor = "#FFFFFF" then
             xTrColor = "#F0F0F0"
          ELSE
             xTrColor = "#FFFFFF"
          END IF

          rs.movenext
       loop

       Tabla = "<BR>" & ofv.GenTabla("","AAA BGCOLOR='#f0f0f0' ","","99%","","0","","1","2","","",Filan,"si")

     END IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''  POR SALVAGUARDAS MITIGANTES ''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  ELSE

'''''''''''''''''''''''''''''
'' POR AMENAZA ASOCIADA
'''''''''''''''''''''''''''''

    sImageSEM = "<font face='corflaticon' size=2 color='#xxxxxx' style='border-style:solid;border-width:1;border-color:#f0f0f0;' >&#0110;</font>"

    sForm = "formPER" & RS(0)

    WKCODRIESGO = "NO"


    Celda  = Celda & ofv.GenCelda("","uti colspan=4  style='background-color:#336699;color:#ffffff;'  ","","","","","","CALIFICACION DE LA AMENAZA ASOCIADA","si")

    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


       WKCODRIESGO = "SI"

       IF WKCODRIESGO = "NO" THEN

          wspcolor0AI  = wspcolor4
          wspcolor0AP  = wspcolor4

          wspcolor0A   = wspcolor4
          wspcolor0B   = wspcolor4
          wspcolor0BB  = wspcolor4

          wspDATA      = "---"
          wspDATB      = "---"
          wspDATBB     = "---"

             wspmontobruto    = "0"
             wspmontoresidual = "0"
             wspmontoresidaj  = "0"
             wspDATAdsc       = "-----"
             wspDATArrdsc     = "-----"
             wspDATArradsc    = "-----"
             wspDATprodsc     = "-----"
             wspDATimpdsc     = "-----"

       ELSE

          sRECALC = request.querystring("RECALC")
     '     RESPONSE.WRITE sRECALC


             wspDATA      = rs(14)
             wspDATB      = rs(16)
             wspDATBB     = rs(20)

             wspDATP      = 0
             wspDATI      = 0

             wspDATPO     = RS(9)
             wspDATIO     = RS(10)

             wspDATM      = 0
             wspDATMC     = 0

             wspDATMB     = 0
             wspDATMCB    = 0

             wspDATMBe    = 0
             wspDATMe     = 0

             wspmontobruto    = "0"
             wspmontoresidual = "0"
             wspmontoresidaj  = "0"
             wspDATAdsc       = "-----"
             wspDATArrdsc     = "-----"
             wspDATArradsc    = "-----"
             wspDATprodsc     = "-----"
             wspDATimpdsc     = "-----"

             WKTCONFIDENP = CINT("0" & RS(11))
             WKTINTEGRIDP = CINT("0" & RS(12))
             WKTDISPONIBP = CINT("0" & RS(13))


             WKTIMPACTO = WKTCONFIDENP


             IF WKTIMPACTO < WKTINTEGRIDP THEN WKTIMPACTO = WKTINTEGRIDP
             IF WKTIMPACTO < WKTDISPONIBP THEN WKTIMPACTO = WKTDISPONIBP

             wspporc = 100

             Strsql  = "Select VALOR,porcentaje,obs1 From ORACTPROBOCURRENCIA where OBJID = 0" & wspDATPO
            ' response.write strsql
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if  NOT rs1.eof then
                  wspDATPX    = "0" & RS1(0)
                  wspDATP     = CINT(wspDATPX)
                  wspporc     = CINT("0" & RS1(1))
                  wspDATprodsc     = "&nbsp;" & RS1(1) & "&nbsp;%&nbsp-&nbsp;" & RS1(2)
             END IF

             wspmonto = 0


          IF SRECALC = "SI" THEN

''''''''' ANALIZA EL IMPACTO POR CADA DIMENSION PONDERADO POR CADA AMENAZA (DEGRADACION)



             wspDATI     = 0
             wspDATIX    = "0"
             wspmonto    = WKTIMPACTO

          '   Strsql  = "Select VALOR,impdesde,imphasta, OBJID From ORACTIMPACTO where "
          '   strsql = strsql & " NOT (impdesde > " & WKTIMPACTO & ")  "
         '''   response.write strsql & "<br>" & chr(13)
          '   set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          '   if NOT rs1.eof then
          '      wspDATIO     = RS1(3)
          '   END IF




             SVTOPE = 0
             SVCODI = 0


             Strsql = "Select impdesde, OBJID From ORACTIMPACTO ORDER BY impdesde"
        '''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             DO UNTIL rs1.eof 

                 IF cint(WKTIMPACTO) = 0  THEN
                    wspDATIO     = RS1(1)
                    EXIT DO
                 END IF

                 IF (NOT WKTIMPACTO < SVTOPE) AND (WKTIMPACTO < CDBL("0" & RS1(0))) THEN
                    wspDATIO    = SVCODI
                    EXIT DO
                 END IF

                 SVTOPE = RS1(0)
                 SVCODI = RS1(1)


             RS1.MOVENEXT
             LOOP 




          END IF




             wspDEGRADA = WKTIMPACTO




             Strsql  = "Select VALOR,impdesde,imphasta,obs1,DESCRI From ORACTIMPACTO where OBJID = 0" & wspDATIO
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if  NOT rs1.eof then
                  wspDATIX    = "0" & RS1(0)
                  wspDATI     = CINT(wspDATIX)
                  wspDEGRADA  = Clng("0" & RS1(1))
                  wspDATimpdsc     = RS1(4) & "&nbsp;/&nbsp;HASTA&nbsp;" & wspDEGRADA & "%&nbsp;DE&nbsp;DEGRADACION&nbsp;-&nbsp;" & RS1(3)
             END IF


             wspmonto    = clng("0" & sACTVALUACION)


            if wspDATP > 0 and wspDATI > 0 then

               wspDATM = 1
               wspDATMB = 1
               wspDATMBCN =0

             Strsql = "Select A.* From oractivrcontroles A "
             strsql = strsql & "  where a.periodo = 0" & session("ORPERIODO") & " "
             strsql = strsql & "  AND a.actividad = '" & SACTIVIDAD & "' AND A.CODRIESGO =  '" & RS(7) & "' "
             strsql = strsql & "  ORDER BY A.TIPO "

          ' response.write strsql & "<br>"
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rs1.eof then
                do until rs1.eof





                   Strsql  = "Select VALORD From RGACTEFICIENCIASV where OBJID = 0" & RS1(8)
                   set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                   if NOT rs2.eof then
                      wspDATmX    = "0" & RS2(0)
                      wspDATmt     = CINT(wspDATmX)

                    IF RS1("TIPO") = 1 or (RS1("TIPO") = 0 and (rs1("SECUEN") & "") <> "0" ) THEN
                      wspDATMCB    = wspDATMCB + 1
                      wspDATMBe     = wspDATMBe + wspDATmt
                      IF RS1("TIPO") = 1 then wspDATMBCN     =  wspDATMBCN + 1
                    END IF

                    IF RS1("TIPO") = 0 THEN
                      wspDATMC    = wspDATMC + 1
                      wspDATMe     = wspDATMe + wspDATmt
                    END IF



                   END IF



                   rs1.movenext
                loop

                if wspDATMC > 0 and wspDATMe > 0 then
                    wspDATM = cdbl(round(wspDATMe / wspDATMC,2))
                else
                    wspDATM = cdbl(0)
                end if

                if (wspDATMCB > 0 and wspDATMBe > 0)  then
                    wspDATMB = cdbl(round((wspDATMBe) / (wspDATMCB),2))
                else
                    wspDATMB = cdbl(0)
                end if

             end if

             SVTOPE = 0
             SVCODI = 0
             SVDSC  = 0
             wspDATMdsc    = "NO"


             Strsql = "Select DESCRIPCION,VALORD,OBJID From RGACTEFICIENCIASV ORDER BY VALORD"
        '''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             DO UNTIL rs1.eof 

                 IF cint(wspDATM) = 0 AND SVCODI > 0 THEN
                    wspDATMdsc    = SVDSC   'DESCRIPCION
                    EXIT DO
                 END IF

                 IF (NOT wspDATM < SVTOPE) AND (wspDATM < CDBL("0" & RS1(1))) THEN
                    wspDATMdsc    = SVDSC   'DESCRIPCION
                    EXIT DO
                 END IF

                 SVTOPE = RS1(1)
                 SVCODI = RS1(2)
                 SVDSC  = RS1(0)

             RS1.MOVENEXT
             LOOP 

             IF wspDATMdsc = "NO" THEN
                IF SVCODI = 0 THEN
                   wspDATMdsc    = "NO ESPECIFICADO"
                ELSE
                   wspDATMdsc    = SVDSC
                END IF

             END IF 



             SVTOPE = 0
             SVCODI = 0
             SVDSC  = 0
             wspDATMBdsc    = "NO"


             Strsql = "Select DESCRIPCION,VALORD,OBJID From RGACTEFICIENCIASV ORDER BY VALORD"
        '''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             DO UNTIL rs1.eof 

                 IF cint(wspDATMB) = 0 AND SVCODI > 0 THEN
                    wspDATMBdsc    = SVDSC   'DESCRIPCION
                    EXIT DO
                 END IF

                 IF (NOT wspDATMB < SVTOPE) AND (wspDATMB < CDBL("0" & RS1(1))) THEN
                    wspDATMBdsc    = SVDSC   'DESCRIPCION
                    EXIT DO
                 END IF

                 SVTOPE = RS1(1)
                 SVCODI = RS1(2)
                 SVDSC  = RS1(0)

             RS1.MOVENEXT
             LOOP 

             IF wspDATMBdsc = "NO" THEN
                IF SVCODI = 0 THEN
                   wspDATMBdsc    = "NO ESPECIFICADO"
                ELSE
                   wspDATMBdsc    = SVDSC
                END IF

             END IF 

             
             wspDATM = (100 - wspDATM)
             wspDATMB = (100 - wspDATMB)            




             wspmontobruto = clng(clng(wspmonto * wspporc / 100) * wspDEGRADA / 100)


             wspDATA      = wspDATP * wspDATI

          ''   wspDATB      = cdbl(round(wspDATA / wspDATM,3))
          ''   wspDATBB     = cdbl(round(wspDATA / wspDATMB,3))

             wspDATB      = cdbl(round(wspDATA * wspDATM / 100,3))
             wspDATBB     = cdbl(round(wspDATA * wspDATMB / 100,3))


           ''  wspmontoresidual = clng(wspmontobruto / wspDATM)
            '' wspmontoresidaj  = clng(wspmontobruto / wspDATMB)


             wspmontoresidual = clng(wspmontobruto * wspDATM / 100)
             wspmontoresidaj  = clng(wspmontobruto * wspDATMB / 100)



             WSPRECALC = ""

             WSPNEWRB =  clng(wspporc * wspDEGRADA  / 100)


             SVTOPE = 0
             SVCODI = 0
             SVDSC  = 0
             WSPNEWRBdsc    = "NO"


             Strsql = "Select * From RGACTRANGORIESGO ORDER BY PORCENTAJE"
        '''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             DO UNTIL rs1.eof 

                 IF cint(WSPNEWRB) = 0  THEN
                    SVTOPE = RS1(4)
                    SVCODI = RS1(0)
                    SVDSC  = RS1(2)
                    SVCOL  = RS1(3)
                    WSPNEWRBdsc    = SVDSC   'DESCRIPCION
                    wspDATAdsc     = SVDSC   'DESCRIPCION
                    wspDATAdsccl   = SVCOL   'COLOR
                    EXIT DO
                 END IF

                 IF (NOT WSPNEWRB < SVTOPE) AND (WSPNEWRB < CDBL("0" & RS1(4))) THEN
                    SVTOPE = RS1(4)
                    SVCODI = RS1(0)
                    SVDSC  = RS1(2)
                    SVCOL  = RS1(3)
                    WSPNEWRBdsc    = SVDSC   'DESCRIPCION
                    wspDATAdsc     = SVDSC   'DESCRIPCION
                    wspDATAdsccl   = SVCOL   'COLOR
                    EXIT DO
                 END IF

                    SVTOPE = RS1(4)
                    SVCODI = RS1(0)
                    SVDSC  = RS1(2)
                    SVCOL  = RS1(3)

             RS1.MOVENEXT
             LOOP 

             IF WSPNEWRBdsc = "NO" THEN
                IF SVCODI = 0 THEN
                   WSPNEWRBdsc    = "NO ESPECIFICADO"
                ELSE
                    WSPNEWRBdsc    = SVDSC   'DESCRIPCION
                    wspDATAdsc     = SVDSC   'DESCRIPCION
                    wspDATAdsccl   = SVCOL   'COLOR
                END IF

             END IF 
            






         '    Strsql  = "Select descri,REQPLAN From ORACTriesgo where "
         '    Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
         '    Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "
  ' response.write wspDATMBCN & "-----"
         '    NIVEL=-1
         '    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         '    if NOT rs1.eof then

         '       FOR II=1 TO TTABRANRIESGO

'                    IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
'                    ' response.write RS1(0)&"--" & rs1(1)  & "<br>" & chr(13)
'                       wspDATArrdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
'                       wspDATArrdsccl    = TABRANRIESGOS(II,3) 'COLOR
'                       NIVEL=CINT("0" & RS1(0))''

'
'                      IF RS1(1) = "SI" THEN
'                          WSPRECALC = " ,REQUIEREPLAN='SI' "
'                          SREQUIEREPLAN="SI"
'
'                       else
'                     ' WSPRECALC = " ,REQUIEREPLAN='NO' "
'                       SREQUIEREPLAN="NO"'

'                         if wspDATMBCN  <> 0 then
'                           WSPRECALC = ""
'                         else
'                          WSPRECALC = " ,REQUIEREPLAN='NO' "
'                         end if
'                       end if
'                       EXIT FOR
'                    END IF
'
'                NEXT
'
'              ''  wspDATArrdsc    = RS1(0)
'             END IF

'             Strsql  = "Select descri From ORACTriesgo where "
'             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
'             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "

          ''   Strsql  = "Select descri From ORNIVriesgo where valorhasta > 0" & replace(cstr(wspDATBB & ""),",",".")
 ' response.write strsql & "<br>" & chr(13)
'             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            ' NIVEL=-1
'             if NOT rs1.eof then

'                FOR II=1 TO TTABRANRIESGO

'                    IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
'                       wspDATArrAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
'                       wspDATArrAdsccl    = TABRANRIESGOS(II,3) 'COLOR
'                      ' NIVEL=CINT("0" & RS1(0))
'                       EXIT FOR
'                    END IF

'                NEXT


'           ''     wspDATArradsc    = RS1(0)
'             END IF




             Strsql  = "Select * From ORACTriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "
  ' response.write wspDATMBCN & "-----"
             NIVEL=-1
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then

                       wspDATArrdsc      = RS1(1) 'DESCRIPCION
                       wspDATArrdsccl    = RS1(5) 'COLOR
                       NIVEL             = CINT("0" & RS1(0))

                       IF RS1(4) = "SI" THEN
                          WSPRECALC = " ,REQUIEREPLAN='SI' "
                          SREQUIEREPLAN="SI"

                       else

                          SREQUIEREPLAN="NO"

                         if wspDATMBCN  <> 0 then
                            WSPRECALC = ""
                         else
                           WSPRECALC = " ,REQUIEREPLAN='NO' "
                         end if

                      end if


             END IF


             Strsql  = "Select * From ORACTriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "
 ' response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             if NOT rs1.eof then
                       wspDATArrAdsc     = RS1(1)   'DESCRIPCION
                       wspDATArrAdsccl   = RS1(5) 'COLOR

             END IF



                     ''  wspmontoresidual    = wspmontoresidual & "-" & wspDATB & "-" & wspDATM
                     ''  wspmontoresidaj    = wspmontoresidaj & "-" & wspDATBB & "-" & wspDATMB


          end if
'response.write nivel

          IF SRECALC = "SI" THEN



            if wspDATP > 0 and wspDATI > 0 then



               Strsql = "UPDATE RGACTANALISIS SET IMPACTO = " & wspDATIO & ",RIESGOBRUTO = " & wspDATA & ", RIESGORESIDUAL = " & replace(cstr(wspDATB & ""),",",".") & ", RIESGORESIDUALNV = " & replace(cstr(wspDATBB & ""),",",".") & WSPRECALC
               Strsql = Strsql & " where CODRIESGO = '" & sRIESGO & "' "

 ' response.write strsql & "<br>" & chr(13)
               CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

               Strsql = "Select A.*, b.actividad AS A1, c.actividad AS A2 From RGACTANALISIS A "
               strsql = strsql & " left join (select distinct actividad, CODRIESGO "
               strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 0) b "
               strsql = strsql & " on (a.actividad = b.actividad AND A.CODRIESGO = B.CODRIESGO )"
               strsql = strsql & " left join (select distinct actividad, CODRIESGO "
               strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 1) c "
               strsql = strsql & " on (a.actividad = c.actividad AND A.CODRIESGO = c.CODRIESGO )"
               strsql = strsql & " where A.CODRIESGO =  '" & sRIESGO & "' "
           ''    response.write strsql & "<br>" & chr(13)
               set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             end if

          END IF








       END IF


    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","xxxxxO","si")

    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

    sInput = wspDATp & " - " & wspDATI & " / " & wspDATpX & " - " & wspDATIX
    Celda  = Celda & ofv.GenCelda("","UTI COLSPAN=2 ","","","","","",sInput,"si")
  '  FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","AMENAZAS","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")


    Strsql = " SELECT A.*,B.ACTFAMILYCOD FROM ORAMENAZAS A "
    Strsql = Strsql & " INNER JOIN RGACTFAMILYRSK B ON (A.OBJID = B.ACTRSKCOD) "
    Strsql = Strsql & " WHERE B.ACTFAMILYCOD = 0" & sACTFAMILY & " ORDER BY A.DESCRI"

    sInput = ofv.GenerarCombo(sForm,RS(8).NAME,RS(8).value,Strsql,cn,0,2,"")
    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","DESCRIPCION ADICIONAL","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

        nCol = 100

        if  len(wdescri) = 0 then wdescri = " "

        if  not  isnull(rs(16)) then
            wdescri = rs(16) & ""
        end if

        t3b = "<font class=at><textarea class=dt "
        t3b = t3b & "name=" & rs(16).name & " cols=" & nCol & " rows=2 "

        t3b = t3b & " onblur=" & sForm & ".submit() "
        t3b = t3b & " onkeypress=" & chr(34)
        t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
        t3b = t3b & " extenso'); return false;}" & chr(34) & " >"

        xdatos =  wdescri
        t3b    = t3b & xdatos & "</textarea></font>"
        Celda  = Celda & ofv.GenCelda("","UT colspan=2","","100%","","","",t3b,"si")
        FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","VALOR ACTIVO","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
    sInput = "$&nbsp;" & sACTVALUACION & "."
    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


       wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0AP & ";' "

       wsstyloF2   = " style='border-style:SOLID;border-width:2;"
       wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0AP & ";' "



    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","20%","","","","PROBABILIDAD","si")
    Celda  = Celda & ofv.GenCelda("","UT ","","2%","","","","&nbsp;","si")
    Strsql = "Select * From ORACTPROBOCURRENCIA ORDER BY VALOR"
    sInput = ofv.GenerarCombo(sForm,RS(9).NAME,RS(9).value,Strsql,cn,0,2,"")
    Celda  = Celda & ofv.GenCelda("","UT  ","","20%","","","",sInput,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","58%","","","","&nbsp;" & wspDATprodsc,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

       wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0AI & ";' "

       wsstyloF2   = " style='border-style:SOLID;border-width:2;"
       wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0AI & ";' "



      wscomK = "0" & chr(9) & "10" & chr(9) & "20" & chr(9) & "30"
      wscomK = wscomK & chr(9) & "40" & chr(9) & "50" & chr(9) & "60"
      wscomK = wscomK & chr(9) & "70" & chr(9) & "80" & chr(9) & "90" & chr(9) & "100"


      wscomD = "N/A" & chr(9) & "10%" & chr(9) & "20%" & chr(9) & "30%"
      wscomD = wscomD & chr(9) & "40%" & chr(9) & "50%" & chr(9) & "60%"
      wscomD = wscomD & chr(9) & "70%" & chr(9) & "80%" & chr(9) & "90%" & chr(9) & "100%"


    FilaT2 = ""

    Celda2  = ""
    Celda2 = Celda2 & ofv.gencelda("","UTI COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",wspDATimpdsc,"si")
    FilaT2 = filaT2 & ofv.GenRow("","","","","25","","",Celda2,"si")

    Celda2  = ""
    Celda2 = Celda2 & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#5599CC;border-style:SOLID;border-width:1;border-color:#000000;' ","XCENTER","33%","","","","CONFIDENCIALIDAD&nbsp;","si")
    Celda2 = Celda2 & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#5599CC;border-style:SOLID;border-width:1;border-color:#000000;' ","XCENTER","33%","","","","INTEGRIDAD&nbsp;","si")
    Celda2 = Celda2 & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#5599CC;border-style:SOLID;border-width:1;border-color:#000000;' ","XCENTER","33%","","","","DISPONIBILIDAD&nbsp;","si")

    FilaT2 = filaT2 & ofv.GenRow("","","","","25","","",Celda2,"si")

    Celda2  = ""
 
    WKTCONFIDEN = CINT("0" & RS(11))
    WKTINTEGRID = CINT("0" & RS(12))
    WKTDISPONIB = CINT("0" & RS(13))

    SINPUTD1 = ofv.gencomboFX(sForm,RS(11).NAME,WKTCONFIDEN,wscomK,wscomD," ")
    SINPUTD2 = ofv.gencomboFX(sForm,RS(12).NAME,WKTINTEGRID,wscomK,wscomD," ")
    SINPUTD3 = ofv.gencomboFX(sForm,RS(13).NAME,WKTDISPONIB,wscomK,wscomD," ")

    Celda2 = Celda2 & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","",SINPUTD1,"si")
    Celda2 = Celda2 & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","",SINPUTD2,"si")
    Celda2 = Celda2 & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","",SINPUTD3,"si")

    FilaT2 = filaT2 & ofv.GenRow("","","","","25","","",Celda2,"si")

    WSPTABIMPACTO = ofv.GenTabla("","AAA  ","","100% ","","0","","0","0","","",FilaT2,"si")


    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","IMPACTO","si")
    Celda  = Celda & ofv.GenCelda("","UT ","","","","","","&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2  ","CENTER","","","","",WSPTABIMPACTO,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")











       wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0A & ";' "

       wsstyloF2   = " style='border-style:SOLID;border-width:1;"
       wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0A & ";background-color:#F0F0F0;' "


    IF WSPSIGUE = "SI" THEN

    scomK = "SI" & chr(9) & "NO"
    scomD = "SI" & chr(9) & "NO"


    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","RIESGO REPUTACIONAL","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
    sInput = ofv.gencomboFX(sForm,rs(21).name,rs(21),scomK,scomD," ")
    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


    IF (RS(21) & "") = "SI" THEN

       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","TT","","","","","","CALIF.R.REPUTACIONAL","si")
       Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

       Strsql = "Select * From ORRREPUTACIONAL ORDER BY DESCRI"
       sInput = ofv.GenerarCombo(sForm,RS(22).NAME,RS(22).value,Strsql,cn,0,1,"")
       Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


    end if

  END IF





    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","RIESGO INHERENTE","si")



    Celda  = Celda & ofv.GenCelda("","UT ","","","","","","&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Vt " & wsstyloF2,"CENTER","","","","",wspDATAdsc,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;$&nbsp;" & wspmontobruto,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



          Form1N     = "RGACTCIRCUITO" & rs(0)

          sAccion    = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO" & "&CONTROLTIPO=0"
          sAccion    = sAccion & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster & "&CONTROLES=SI"
          Form1      = ofv.genform(Form1N,"AAA STYLE='display:none;' ",sAccion,"","post","","<P>&nbsp;</P>","si")


          sImageO0   = "4"
          wspcolor0C  = "#000099"

              Strsqlx = "Select A.* From oractivrcontroles A "
		  	  strsqlx = strsqlx & "  where a.califmitigante <>0  and tipo=0 and a.periodo = 0" & session("ORPERIODO") & " "
		  	  strsqlx = strsqlx & " AND a.actividad = '" & SACTIVIDAD & "' AND A.CODRIESGO =  '" & RS(7) & "' "

		       set rs1rd = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
		       sread=""
		       if rs1rd.eof then
		        sread=" readonly "
               end if
               call ofv.cerrarconsulta(rs1rd)
''RESPONSE.WRITE strsqlx
           Strsqlx = "Select A.* From oractivrcontroles A "
	   	   strsqlx = strsqlx & "  where a.califmitigante<>0 and tipo=1 and a.periodo = 0" & session("ORPERIODO") & " "
	   	   strsqlx = strsqlx & " AND a.actividad = '" & SACTIVIDAD & "' AND A.CODRIESGO =  '" & RS(7) & "' "
	   		  	 '  RESPONSE.WRITE strsqlx
	   	   set rs1rd = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
	   	    sreadCN="NO"
		    if NOT rs1rd.eof then
		   	 sreadCN="SI"
            end if

	   	   call ofv.cerrarconsulta(rs1rd)

          if  (RS("A1")& "") ="" or sread <> "" then

             sImageO0   = sImageO1
             wspcolor0C  = wspcolor1

          else

             sImageO0   = sImageO2
             wspcolor0C  = wspcolor2

          end if


        wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;"
        wsstyloF   = wsstyloF & "background-color:" & wspcolor0C & ";' "

        sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >" & sImageO0 & "</font>"



    Celda  = ""

    Celda  = Celda & ofv.GenCelda("","TT","","","","","","SALVAGUARDAS","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","ut onclick='document." & Form1N & ".submit();' " & wsstyloF,"","10%","","","",sImageOF  ,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;" & wspDATMdsc,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")




          if (rs(23) & "") = "" OR (RS("A1")& "") ="" or sread <> "" then

             sImageO0   = sImageO1
             wspcolor0C  = wspcolor1

          else

             sImageO0   = sImageO2
             wspcolor0C  = wspcolor2

          end if



       wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0B & ";' "

       wsstyloF2   = " style='border-style:SOLID;border-width:1;"
       wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0B & ";background-color:#F0F0F0;' "

    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","RIESGO RESIDUAL","si")
    Celda  = Celda & ofv.GenCelda("","UT ","","","","","","&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Vt " & wsstyloF2,"CENTER","10%","","","",wspDATArrdsc,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;$&nbsp;" & wspmontoresidual,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

   wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
   wsstyloF   = wsstyloF & "background-color:" & wspcolor0BB & ";' "

   wsstyloF2   = " style='border-style:SOLID;border-width:1;"
   wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0BB & ";background-color:#F0F0F0;' "



    scomK = "SI" & chr(9) & "NO"
    scomD = "SI" & chr(9) & "NO"





    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","REQUIERE PLAN" ,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

 '   if sread="" AND SREQUIEREPLAN="NO" AND sreadCN="NO" then
    if sread="" AND SREQUIEREPLAN="NO"  then

   sInput = ofv.gencomboFX("frmasgx",rs(17).name  ,rs(17),scomK,scomD," ")
     sAccion = "asp/actualizar.asp?fname=" & rs(0)
     wkproccodX = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&CONTROLTIPO=" & SCONTROLTIPO
     wkproccodX = wkproccodX & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster

      VUELTAX = "../" & REPLACE(wkproccodX,"&","_")
      sAccion = sAccion & "&tabla=ORACTIVRIESGO&volver=" & vueltaX
      ' sAccion = sAccion & "&Opcion=SI" & "&CONTROLTIPO=" & SCONTROLTIPO
       Form14  = ofv.genform("frmasgx","AAA STYLE='display:none;' ",sAccion,"","post","",sInput,"si")
    else
       sinput=rs(17)
    end if

    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

    IF WSPSIGUE = "SI" THEN


    'ASIGNAR INDICADOR

   wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#888888;COLOR:#888888;"
   Celda  = ""
	Celda  = Celda & ofv.GenCelda("","TT","","","","","","REQUIERE INDICADOR","si")
	Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
	'Celda  = Celda & ofv.GenCelda("","UT","","","","","",sinput,"si")
	sImageDAP  = "<font face='corflaticon' size=3 color='#FFFFFF'>3</font>"
	' response.write NIVEL

    if wspDATP > 0 and wspDATI > 0 then
	STRSQL="SELECT * FROM ORRELAINDIRIESGO WHERE idindicador is not null and CODRIESGO='" & sRIESGO & "'"

	SET RSRIESGO=ofv.crearconsultaEx(StrSql,cn,1,parametros)

	IF NOT RSRIESGO.EOF THEN
	 wspcolor0C="#009900"
	ELSE
	 wspcolor0C="#990000"
	END IF

	CALL OFV.CERRARCONSULTA(RSRIESGO)
   ' response.write nivel

	IF NIVEL >= 3 THEN
     sInput2 = "SI"
     strsql="update oractivriesgo set requiereindicador='SI' WHERE OBJID=" & RS(0)
     call ofv.crearconsultaEx(strsql,cn,1,parametros)

   else
      sinput2="NO"
      strsql="update oractivriesgo set requiereindicador='NO' WHERE OBJID="& RS(0)
      call ofv.crearconsultaEx(strsql ,cn,1,parametros)
   end if


     Form1N2     = "ORASIGNAINDICADOR" & rs(0)

      if sread="" then
	  sAccion    = "ORASIGNAINDICADOR.asp?opcion=LISTAR&ID=" & rs(0) & "&RIESGO=" & sRIESGO & "&ACTIVIDAD=" & SACTIVIDAD & "&CODMASTER=" &SCODMASTER & "&TIPMASTER=" & STIPMASTER
	  sAccion = sAccion  & "&CantRegMostrar=" & CantRegMostrarx & "&volver="  & volver2
      else
       saccion=""
      end if
	  Form13      = ofv.genform("frmasg","AAA STYLE='display:none;' ",sAccion,"","post","","<P>&nbsp;</P>","si")


    ' response.write wspcolor0C
	  sImageO0   = sImageO2
	  spcolor0C  = wspcolor2
	  wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;"
	  wsstyloF   = wsstyloF & "background-color:" & wspcolor0C & ";' "
      sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >" & sImageO0 & "</font>"

    celdax=ofv.GenCelda("","UT","","5%","","","",sinput2,"si")
    celdax=celdax & ofv.GenCelda("","ut onclick='document." & "frmasg.submit();' " & wsstyloF,"","95%","","","", sImageOF,"si")
    filax= ofv.GenRow("","","","","25","","",Celdax ,"si")



     tablaz=ofv.GenTabla("","AAA  ","","100%","","0","","","","","",Filax,"si")
    else
    tablaz="&nbsp;"
    end if
	Celda  = Celda & ofv.GenCelda("","UT ","","","","","",tablaz,"si")




    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda ,"si")


   END IF



    IF (RS(17) & "") = "SI" THEN

        Celda  = ""
        Celda  = Celda & ofv.GenCelda("","uti colspan=4 ","","","","","","OBSERVACIONES","si")
        FilaN = filaN & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

        nCol = 100

        if  len(wdescri) = 0 then wdescri = "Sin observaciones."

        if  not  isnull(rs(18)) then
            wdescri = rs(18)
        else
         wdescri=""
        end if

        t3b = "<font class=at><textarea class=dt "
        t3b = t3b & "name=" & rs(18).name & " cols=" & nCol & " rows=3 "

        t3b = t3b & " onblur=" & sForm & ".submit() "
        t3b = t3b & " onkeypress=" & chr(34)
        t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
        t3b = t3b & " extenso'); return false;}" & chr(34) & " >"

        xdatos =  wdescri
        t3b    = t3b & xdatos & "</textarea></font>"
        Celda  = ofv.GenCelda("","UT colspan=4","center","100%","","","",t3b,"si")
        FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


        Form2N     = "RGACTCIRCUITONV" & rs(0)
        if sread="" then
         sAccion    = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO" & "&CONTROLTIPO=1"
         sAccion    = sAccion & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster & "&CONTROLES=SI"
        else
         saccion=""
        end if
        Form2      = ofv.genform(Form2N,"AAA STYLE='display:none;' ",sAccion,"","post","","<P>&nbsp;</P>","si")


                 sImageO0   = "4"
	             wspcolor0C  = "#000099"

	                Strsqlx = "Select A.* From oractivrcontroles A "
	   		  	    strsqlx = strsqlx & "  where a.califmitigante<>0 and tipo=1 and a.periodo = 0" & session("ORPERIODO") & " "
	   		  	    strsqlx = strsqlx & " AND a.actividad = '" & SACTIVIDAD & "' AND A.CODRIESGO =  '" & RS(7) & "' "
	   		  	 '  RESPONSE.WRITE strsqlx
	   		        set rs1rd = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
	   		       sread=""
	   		       if rs1rd.eof then
	   		        sread=" readonly "
	               end if

                   call ofv.cerrarconsulta(rs1rd)
	             if (RS("A2")& "") ="" or sread <> "" then

	                sImageO0   = sImageO1
	                wspcolor0C  = wspcolor1

	             else

	                sImageO0   = sImageO2
	                wspcolor0C  = wspcolor2

          end if

        wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;"
        wsstyloF   = wsstyloF & "background-color:" & wspcolor0C & ";' "

        sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >" & sImageO0 & "</font>"

        Celda  = ""

        Celda  = Celda & ofv.GenCelda("","TT style='background-color:#336699;' ","","","","","","SALVAGUARDAS NUEVAS","si")
        Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
        Celda  = Celda & ofv.GenCelda("","ut onclick='document." & Form2N & ".submit();' " & wsstyloF,"","10%","","","",sImageOF,"si")

        if sread="" then
         Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp" &  wspDATMBdsc,"si")
        else
         Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;" ,"si")
        end if

        FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

        Celda  = ""
        Celda  = Celda & ofv.GenCelda("","TT style='background-color:#336699;' ","","","","","","RESIDUAL NUEVO","si")
        Celda  = Celda & ofv.GenCelda("","UT ","","","","","","&nbsp;","si")
        Celda  = Celda & ofv.GenCelda("","Vt " & wsstyloF2,"CENTER","","","","",wspDATArradsc,"si")
        Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;$&nbsp;" & wspmontoresidaj,"si")
        FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


    END IF













         IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 )  THEN
            sAccion = "asp/actualizar.asp?fname=" & rs(0)
            sAccion = sAccion & "&tabla=RGACTANALISIS&volver=" & vuelta
         ELSE
            sAccion = ""
         END IF



         FilaN   = ofv.genform(sForm,"",sAccion,"","post","",Filan,"si")





       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF



    Tabla = "<BR><CENTER>" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","95%","","0","","5","2","","",Filan,"si") & "</CENTER>"
    Tabla = Tabla & Form1 & Form2 & Form13 & Form14


   END IF







  end if






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

