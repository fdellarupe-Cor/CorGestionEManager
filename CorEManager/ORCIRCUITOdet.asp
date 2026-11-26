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


 wkproccod = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=SI&CONTROLTIPO=" & SCONTROLTIPO
 'wkproccod = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&CONTROLTIPO=" & SCONTROLTIPO
 wkproccod = wkproccod & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster

 VUELTA = "../" & REPLACE(wkproccod,"&","_")

 wkproccod = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&CONTROLES=SI&CONTROLTIPO=" & SCONTROLTIPO
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

    wkproccod = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RIESGO=" & SRIESGO & "&RECALC=SI&CONTROLTIPO=" & SCONTROLTIPO
    wkproccod = wkproccod & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster

 ELSE

    wkproccod = "ORCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&CONTROLTIPO=" & SCONTROLTIPO
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
      sAccion = sAccion & "&Opcion=SI" & "&CONTROLTIPO=" & SCONTROLTIPO


  else

      WKPROCDESC = "no se encuentra " &  SACTIVIDAD


  END IF






''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA CONTROLES MITIGANTES - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN



        strsql = "Insert into ORACTIVRCONTROLES "
        strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
        strsql = strsql & " SECUEN, CODRIESGO, ITEMCONTROL, CALIFMITIGANTE,TIPO,MODOCONT) "
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & session("ORPERIODO") & "','" & sACTIVIDAD & "','" & sCODPROCED & "',"
        strsql = strsql &  "'" & SCODMASTER & "','" & STIPMASTER & "',"
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



   Strsql = "Select A.*, b.actividad AS A1 ,c.actividad AS A2 From ORACTIVRIESGO A "
   strsql = strsql & " left join (select distinct actividad, CODRIESGO "
   strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 0) b "
   strsql = strsql & " on (a.actividad = b.actividad AND A.CODRIESGO = B.CODRIESGO )"
   strsql = strsql & " left join (select distinct actividad, CODRIESGO "
   strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 1) c "
   strsql = strsql & " on (a.actividad = c.actividad AND A.CODRIESGO = c.CODRIESGO )"
   strsql = strsql & " where A.CODRIESGO =  '" & sRIESGO & "' "

'response.write CHR(13) & strsql & CHR(13)
'response.write CHR(13) & CHR(13)

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if not rs.eof then
  if rs("a1")&"" ="" then
       strsql = "Insert into ORACTIVRCONTROLES "
  	         strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
  	         strsql = strsql & " SECUEN, CODRIESGO, ITEMCONTROL, CALIFMITIGANTE,TIPO,MODOCONT) "
  	         strsql = strsql & " Values ("
  	         strsql = strsql &  "'" & session("ORPERIODO") & "','" & sACTIVIDAD & "','" & sCODPROCED & "',"
  	         strsql = strsql &  "'" & SCODMASTER & "','" & STIPMASTER & "',"
  	         strsql = strsql & "'10','" & SRIESGO & "',"
  	         strsql = strsql & "'0','0','" & SCONTROLTIPO & "','1'"
  	         strsql = strsql & ")"
  	      '   RESPONSE.WRITE STRSQL
          call ofv.crearconsultaEx(StrSql,cn,1,volver)
       end if


  end if

  call ofv.cerrarconsulta(rs)
  Strsql = "Select A.*, b.actividad AS A1 ,c.actividad AS A2 From ORACTIVRIESGO A "
     strsql = strsql & " left join (select distinct actividad, CODRIESGO "
     strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 0) b "
     strsql = strsql & " on (a.actividad = b.actividad AND A.CODRIESGO = B.CODRIESGO )"
     strsql = strsql & " left join (select distinct actividad, CODRIESGO "
     strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 1) c "
     strsql = strsql & " on (a.actividad = c.actividad AND A.CODRIESGO = c.CODRIESGO )"
     strsql = strsql & " where A.CODRIESGO =  '" & sRIESGO & "' "

	'response.write(CHR(13) & CHR(13))
	'response.write strsql
	'response.write(CHR(13) & CHR(13))

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
       Fila = ""

     Celda = ""

   IF wspTIPO = "AI" THEN
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTE LA AMENAZA ASOCIADA AL ACTIVO","si")
   ELSE
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTE EL RIESGO ASOCIADO A LA ACTIVIDAD","si")
   END IF


     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
     Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

     Tabla = "<CENTER><BR><BR><BR><BR>" & Tabla & "</CENTER>"

  else



      sAccion = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&CONTROLES=SI"
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




   Celda = ""
   Celda = Celda & ofv.GenCelda("","uti style='Xbackground-color:#993333;color:#ffffff;' ","","","","","",sInput2 & ":&nbsp;" & WKPROCACCDESC,"si")
   Fila  = ofv.GenRow("","","","","30","","",Celda,"si")
   TablaH2 = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;background-color:#aaaaaa;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")



    xTrColor = "#FFFFFF"

    CeldaK  = ""
    Celda  = ""

  IF SCONTROLES = "SI" THEN

'''''''''''''''''''''''''''''
'' POR CONTROLES MITIGANTES
'''''''''''''''''''''''''''''
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

	'RESPONSE.Write(CHR(13) & CHR(13))
	'response.write strsql
	'RESPONSE.Write(CHR(13) & CHR(13))
	
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if rs.eof then


       Celda = ""

   IF wspTIPO = "AI" THEN
       Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN SALVAGUARDAS MITIGANTES PARA LA AMENAZA ASOCIADA","si")
   else
       Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN CONTROLES MITIGANTES PARA EL RIESGO ASOCIADO","si")
   end if
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



   IF wspTIPO = "AI" THEN

       Celda = ""

       Celda  = Celda & ofv.GenCelda("","uti colspan=5  style='background-color:#336699;color:#ffffff;'  ","","","","","","CALIFICACION DE LA AMENAZA ASOCIADA","si")

       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



       wktitdeleteS = "Confirma la eliminacion de la salvaguarda mitigante???."

 ELSE

       Celda = ""

       Celda  = Celda & ofv.GenCelda("","uti colspan=5  style='background-color:#888888;color:#ffffff;'  ","","","","","","CALIFICACION DEL RIESGO ASOCIADO","si")

       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

       wktitdeleteS = "Confirma la eliminacion del control mitigante???."


 END IF

   IF wspTIPO = "AI" THEN

       Celda = ""


       Celda  = Celda & ofv.GenCelda("","uti colspan=2  style='background-color:#336699;color:#ffffff;'  ","","","","","","SALVAGUARDAS MITIGANTES","si")
       Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#336699;color:#ffffff;'  ","CENTER","","","","","TIPO","si")
       Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#336699;color:#ffffff;'  ","CENTER","","","","","CALIFICACION","si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



 ELSE


       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","uti colspan=2  style='background-color:#888888;color:#ffffff;'  ","","","","","","CONTROLES MITIGANTES","si")
       Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#888888;color:#ffffff;'  ","CENTER","","","","","TIPO","si")
       Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#888888;color:#ffffff;'  ","CENTER","","","","","CALIFICACION","si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


 END IF
	
	'RESPONSE.Write(CHR(13) & CHR(13))
	'response.write strsql
	'RESPONSE.Write(CHR(13) & CHR(13))
	
       do until rs.eof


          Fila   = ""

          sForm = "formPER" & RS(0)

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
       ''   sAccion    = "respproceso.asp?codobjeto=" & rs(0) & "&tipobjeto=" & "RGCM" & ""

          sAccion    = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&OPCION=DEL&CODCONTROL=" & RS(0) & "&CONTROLTIPO=" & SCONTROLTIPO
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


        '  Celda  = Celda & ofv.GenCelda("","ut","CENTER","2","","","",sImageP,"si")

   IF wspTIPO = "AI" THEN

       if wkmodif = "S" then

          Strsql = "Select * From ORSALVAGUARDAS ORDER BY DESCRI"
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

   ELSE



       if wkmodif = "S" then

          Strsql = "Select * From ORCONTROLRIESGO ORDER BY descri"
          sInput = ofv.GenerarCombo(sForm,RS(7).NAME,RS(7).value,Strsql,cn,0,2,"")
       else
          sInput = "-----"

          strsql = "Select descri From ORCONTROLRIESGO where OBJID = 0" & RS(7)
          set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsniv.eof then sInputx = rsniv(0)
          call ofv.cerrarconsulta(rsniv)

          sInput = sInputx

       end if

          Celda  = Celda & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","","70%","","","",sInput,"si")

   END IF


       if wkmodif = "S" then
          'scomK = "1" & chr(9) & "2" & chr(9) & "3"
          'scomD = "AUTOMATICO" & chr(9) & "SEMIAUTOMATICO" & chr(9) & "MANUAL"
          'sInput = ofv.gencomboFX(sForm,rs(11).name,rs(11),scomK,scomD," ")
           Strsql = "Select * From ORTIPOCONTROLES   Order By DESCRI"
           sInput = ofv.GenerarCombo(sForm,RS(11).NAME,RS(11).value,Strsql,cn,0,1,"")
       else
          sInput = "-----"

        '  if (rs(11) & "") = "1" then
        '     sInput = "AUTOMATICO"
        '  elseif (rs(11) & "") = "2" then
        '     sInput = "SEMIAUTOMATICO"
        '  elseif (rs(11) & "") = "3" then
         '    sInput = "MANUAL"
        '  end if
          Strsql = "Select * From ORTIPOCONTROLES WHERE OBJID=0" & RS(11)
          set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			if not rsniv.eof then sInput = rsniv(1)
			call ofv.cerrarconsulta(rsniv)


       'end if

       end if



          Celda  = Celda & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")


       if wkmodif = "S" then
          Strsql = "Select * From ORNIVEFICIENCIA ORDER BY VALOR"
          sInput = ofv.GenerarCombo(sForm,RS(8).NAME,RS(8).value,Strsql,cn,0,4,"")
       else
          sInput = "-----"

          strsql = "Select descri From ORNIVEFICIENCIA where OBJID = 0" & RS(8)
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
  ELSE

'''''''''''''''''''''''''''''
'' POR RIESGO ASOCIADO
'''''''''''''''''''''''''''''

 sImageSEM = "<font face='corflaticon' size=2 color='#xxxxxx' style='border-style:solid;border-width:1;border-color:#f0f0f0;' >&#0110;</font>"

   sForm = "formPER" & RS(0)

    WKCODRIESGO = "NO"

   IF wspTIPO = "AI" THEN
    Celda  = Celda & ofv.GenCelda("","uti colspan=4  style='background-color:#336699;color:#ffffff;'  ","","","","","","CALIFICACION DE LA AMENAZA ASOCIADA","si")

   ELSE
    Celda  = Celda & ofv.GenCelda("","uti colspan=4  style='background-color:#888888;color:#ffffff;'  ","","","","","","CALIFICACION DEL RIESGO ASOCIADO","si")

   END IF

    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



    'IF RS(8) > 0 THEN WKCODRIESGO = "SI"

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


             wspDATA      = rs(11)
             wspDATB      = rs(12)
             wspDATBB     = rs(17)

             wspDATP      = 0
             wspDATI      = 0

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



             wspporc = 100

             Strsql  = "Select VALOR,porcentaje,obs1,obs2 From ORNIVPROBOCURRENCIA where OBJID = 0" & RS(9)
            ' response.write strsql
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then
                wspDATPX    = "0" & RS1(0)
                wspDATP     = CINT(wspDATPX)
                wspporc     = CINT("0" & RS1(1))
                wspDATprodsc     = "&nbsp;" & RS1(1) & "&nbsp;%&nbsp-&nbsp;" & RS1(2) & "&nbsp;%&nbsp-&nbsp;" & RS1(3) 
             END IF

             wspmonto = 0

             Strsql  = "Select VALOR,impdesde,imphasta,obs1 From ORNIVIMPACTO where OBJID = 0" & RS(10)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then
                wspDATIX    = "0" & RS1(0)
                wspDATI     = CINT(wspDATIX)
                wspmonto    = clng((Clng("0" & RS1(1)) + Clng("0" & RS1(2))) / 2)
                wspDATimpdsc     = "&nbsp;$&nbsp;" & RS1(1) & " / $ " & RS1(2) & "&nbsp;-&nbsp;" & RS1(3)
             END IF


            if wspDATP > 0 and wspDATI > 0 then

               wspDATM = 1
               wspDATMB = 1
               wspDATMBCN =0

             Strsql = "Select A.* From oractivrcontroles A "
             strsql = strsql & "  where a.periodo = 0" & session("ORPERIODO") & " "
             strsql = strsql & " AND a.actividad = '" & SACTIVIDAD & "' AND A.CODRIESGO =  '" & RS(7) & "' "
             strsql = strsql & " ORDER BY A.TIPO "

          ' response.write strsql & "<br>"
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rs1.eof then
                do until rs1.eof

                   Strsql  = "Select VALOR From ORNIVeficiencia where OBJID = 0" & RS1(8)
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
                    wspDATM = cdbl(1)
                 end if

                 if (wspDATMCB > 0 and wspDATMBe > 0)  then
                    wspDATMB = cdbl(round((wspDATMBe) / (wspDATMCB),2))
                 else
                    wspDATMB = cdbl(1)
                 end if


            '''     if (wspDATMCB > 0 and wspDATMBe > 0) or (wspDATMC > 0 and wspDATMe > 0) then
            '''        wspDATMB = cdbl(round((wspDATMBe + wspDATMe) / (wspDATMCB + wspDATMC),2))
            '''     else
            '''        wspDATMB = cdbl(1)
            '''     end if


             end if


          Strsql = "Select DESCRI From ORNIVEFICIENCIA WHERE VALOR =0" & cint(wspDATM)
 '''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then

                wspDATMdsc    = RS1(0)   'DESCRIPCION

             END IF


          Strsql = "Select DESCRI From ORNIVEFICIENCIA WHERE VALOR =0" & cint(wspDATMB)
' response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then

                wspDATMBdsc    = RS1(0)   'DESCRIPCION

             END IF




             wspmontobruto = clng(wspmonto * wspporc / 100)


             wspDATA     = wspDATP * wspDATI

             wspDATB      = cdbl(round(wspDATA / wspDATM,3))
             wspDATBB     = cdbl(round(wspDATA / wspDATMB,3))


             wspmontoresidual = clng(wspmontobruto / wspDATM)
             wspmontoresidaj  = clng(wspmontobruto / wspDATMB)

             WSPRECALC = ""


             Strsql  = "Select descri From ORNIVRIESGOBRUTO where NIVPROB =  0" & RS(9) & " AND NIVIMPA = 0" & RS(10) & " "
         ' response.write strsql & "<br>" & chr(13)
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

             Strsql  = "Select descri,REQPLAN,REQINDICADOR From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "
   '''response.write  Strsql
             NIVEL=-1
             SREQINDICADOR="NO"
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             if NOT rs1.eof then

                FOR II=1 TO TTABRANRIESGO
                  '  response.write RS1(0)&"--" & TABRANRIESGOS(II,0)  &"--" &TABRANRIESGOS(II,2) &"<br>" & chr(13)
                    IF RS1(0)  =  TABRANRIESGOS(II,0) THEN

                       wspDATArrdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArrdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       NIVEL=CINT("0" & RS1(0))
                       SREQINDICADOR=RS1("REQINDICADOR")

                      IF RS1(1) = "SI" THEN
                          WSPRECALC = " ,REQUIEREPLAN='SI' "
                          SREQUIEREPLAN="SI"

                       else
                     ' WSPRECALC = " ,REQUIEREPLAN='NO' "
                          SREQUIEREPLAN="NO"

                         if wspDATMBCN  <> 0 then
                           WSPRECALC = ""
                         else
                          WSPRECALC = " ,REQUIEREPLAN='NO' "
                         end if
                       end if
                       EXIT FOR
                    END IF

                NEXT

              ''  wspDATArrdsc    = RS1(0)
             END IF

                 ''''      wspDATArrdsc    = wspDATArrdsc & "/" & wspDATB


             Strsql  = "Select descri From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "

          ''   Strsql  = "Select descri From ORNIVriesgo where valorhasta > 0" & replace(cstr(wspDATBB & ""),",",".")
 ' response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            ' NIVEL=-1
             if NOT rs1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
                       wspDATArrAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArrAdsccl    = TABRANRIESGOS(II,3) 'COLOR
                      ' NIVEL=CINT("0" & RS1(0))
                       EXIT FOR
                    END IF

                NEXT


           ''     wspDATArradsc    = RS1(0)
             END IF


                     ''  wspmontoresidual    = wspmontoresidual & "-" & wspDATB & "-" & wspDATM
                     ''  wspmontoresidaj    = wspmontoresidaj & "-" & wspDATBB & "-" & wspDATMB


          end if
'response.write nivel
          sRECALC = request.querystring("RECALC")
     '     RESPONSE.WRITE sRECALC
          IF SRECALC = "SI" THEN



            if wspDATP > 0 and wspDATI > 0 then



               Strsql = "UPDATE ORACTIVRIESGO SET RIESGOBRUTO = " & wspDATA & ", RIESGORESIDUAL = " & replace(cstr(wspDATB & ""),",",".") & ", RIESGORESIDUALNV = " & replace(cstr(wspDATBB & ""),",",".") & WSPRECALC
               Strsql = Strsql & " where CODRIESGO = '" & sRIESGO & "' "

 ' response.write strsql & "<br>" & chr(13)
               CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

               Strsql = "Select A.*, b.actividad AS A1, c.actividad AS A2 From ORACTIVRIESGO A "
               strsql = strsql & " left join (select distinct actividad, CODRIESGO "
               strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 0) b "
               strsql = strsql & " on (a.actividad = b.actividad AND A.CODRIESGO = B.CODRIESGO )"
               strsql = strsql & " left join (select distinct actividad, CODRIESGO "
               strsql = strsql & " from oractivrcontroles where periodo = 0" & session("ORPERIODO") & " and tipo = 1) c "
               strsql = strsql & " on (a.actividad = c.actividad AND A.CODRIESGO = c.CODRIESGO )"
               strsql = strsql & " where A.CODRIESGO =  '" & sRIESGO & "' "
           '    response.write strsql & "<br>" & chr(13)
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

   IF wspTIPO = "AI" THEN

    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","AMENAZAS","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

    Strsql = "Select * From ORAMENAZAS ORDER BY DESCRI"
    sInput = ofv.GenerarCombo(sForm,RS(8).NAME,RS(8).value,Strsql,cn,0,2,"")
    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



   ELSE

    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","RIESGO","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

    'Strsql = "Select * From ORIMPRIESGO ORDER BY descri"
    'sInput = ofv.GenerarCombo(sForm,RS(8).NAME,RS(8).value,Strsql,cn,0,2,"")
    'Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    'FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")
	
	StrSQL = "SELECT O3.OBJID, O1.DESCRI + ' / ' + O2.DESCRI + ' / ' + O3.DESCRI AS 'DESCRIPCION' "
	StrSQL = StrSQL & "FROM ORTIPOEVENTO O1 LEFT JOIN ORTIPOCATEGO O2 ON O1.OBJID = O2.TIPOEVENTO " 
	StrSQL = StrSQL & "LEFT JOIN ORIMPRIESGO O3 ON O3.CATEGO = O2.OBJID AND O3.AGRUPA = O1.OBJID WHERE O3.OBJID = " & RS(8)
	
	set RS_Descri = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	
	if RS_Descri.eof then

		sImagenP = "&#0232;"
		Style_Flecha = "font-family:corflaticon;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;"
		IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 )  THEN
			sInput = "<a href='SUBCONSULTARIESGO.asp?ID=" & RS(8).VALUE & "&CantRegMover=" & CantRegMoverx
			sInput = sInput  & "&CantRegMostrar=" & CantRegMostrarx & "&ID_Actividad_Riesgo=" & RS(0) & "&eACTIVIDAD=" & sACTIVIDAD & "&eRIESGO=" & sRIESGO & "&ecodmaster=" & scodmaster & "&etipmaster=" & stipmaster
			sInput = sInput  & "' title='Cambio de Riesgo ' >" & "<font color=#888888>" & "Riesgo sin Asociar "  &  "</font>"  & "<font style=" & Style_Flecha & ">" & sImagenP & "</font> </a>"
		ELSE
			sInput = RS_Descri(1)
		END IF
		Celda  = Celda & ofv.GenCelda("","UT colspan=2","","","","","", sInput ,"si")
		FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")
	
	else

		sImagenP = "&#0232;"
		Style_Flecha = "font-family:corflaticon;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;"
		IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 )  THEN
			sInput = "<a href='SUBCONSULTARIESGO.asp?ID=" & RS(8).VALUE & "&CantRegMover=" & CantRegMoverx
			sInput = sInput  & "&CantRegMostrar=" & CantRegMostrarx & "&ID_Actividad_Riesgo=" & RS(0) & "&eACTIVIDAD=" & sACTIVIDAD & "&eRIESGO=" & sRIESGO & "&ecodmaster=" & scodmaster & "&etipmaster=" & stipmaster
			sInput = sInput  & "' title='Cambio de Riesgo ' >" & "<font color=#888888>" & RS_Descri(1) & "</font>"  & "<font style=" & Style_Flecha & ">" & sImagenP & "</font> </a>"
		ELSE
			sInput = RS_Descri(1)
		END IF
		Celda  = Celda & ofv.GenCelda("","UT colspan=2","","","","","", sInput ,"si")
		FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")
	
	end if
	
	call ofv.cerrarconsulta(RS_Descri)

   END IF



    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","DESCRIPCION ADICIONAL","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

        nCol = 100

        if  len(wdescri) = 0 then wdescri = " "

        if  not  isnull(rs(13)) then
            wdescri = rs(13) & ""
        end if

        t3b = "<font class=at><textarea class=dt "
        t3b = t3b & "name=" & rs(13).name & " cols=" & nCol & " rows=2 "

        t3b = t3b & " onblur=" & sForm & ".submit() "
        t3b = t3b & " onkeypress=" & chr(34)
        t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
        t3b = t3b & " extenso'); return false;}" & chr(34) & " >"

        xdatos =  wdescri
        t3b    = t3b & xdatos & "</textarea></font>"
        Celda  = Celda & ofv.GenCelda("","UT colspan=2","","100%","","","",t3b,"si")
        FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","ORDEN","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
    sInput = ofv.GenerarInput(rs(6).name,rs(6).value,"text",sForm,"5","dtn","")
    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
'    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


       wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0AP & ";' "

       wsstyloF2   = " style='border-style:SOLID;border-width:2;"
       wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0AP & ";' "



    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","20%","","","","PROBABILIDAD","si")
    Celda  = Celda & ofv.GenCelda("","UT ","","2%","","","","&nbsp;","si")
    Strsql = "Select * From ORNIVPROBOCURRENCIA ORDER BY VALOR"
    sInput = ofv.GenerarCombo(sForm,RS(9).NAME,RS(9).value,Strsql,cn,0,2,"")
    Celda  = Celda & ofv.GenCelda("","UT  ","","20%","","","",sInput,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","58%","","","","&nbsp;" & wspDATprodsc,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

       wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0AI & ";' "

       wsstyloF2   = " style='border-style:SOLID;border-width:2;"
       wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0AI & ";' "


    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","IMPACTO","si")
    Celda  = Celda & ofv.GenCelda("","UT ","","","","","","&nbsp;","si")
    Strsql = "Select * From ORNIVIMPACTO ORDER BY VALOR"
    sInput = ofv.GenerarCombo(sForm,RS(10).NAME,RS(10).value,Strsql,cn,0,2,"")
    Celda  = Celda & ofv.GenCelda("","UT  ","","","","","",sInput,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;" & wspDATimpdsc,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


       wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0A & ";' "

       wsstyloF2   = " style='border-style:SOLID;border-width:1;"
       wsstyloF2   = wsstyloF2 & "border-color:" & wspcolor0A & ";background-color:#F0F0F0;' "


    scomK = "SI" & chr(9) & "NO"
    scomD = "SI" & chr(9) & "NO"


    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","RIESGO REPUTACIONAL","si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
    sInput = ofv.gencomboFX(sForm,rs(18).name,rs(18),scomK,scomD," ")
    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


    IF (RS(18) & "") = "SI" THEN

       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","TT","","","","","","CALIF.R.REPUTACIONAL","si")
       Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

       Strsql = "Select * From ORRREPUTACIONAL ORDER BY DESCRI"
       sInput = ofv.GenerarCombo(sForm,RS(19).NAME,RS(19).value,Strsql,cn,0,1,"")
       Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


    end if





    Celda  = ""
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","RIESGO INHERENTE","si")



    Celda  = Celda & ofv.GenCelda("","UT ","","","","","","&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Vt " & wsstyloF2,"CENTER","","","","",wspDATAdsc,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;$&nbsp;" & wspmontobruto,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



          Form1N     = "ORCIRCUITO" & rs(0)

          sAccion    = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO" & "&CONTROLTIPO=0"
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
'RESPONSE.WRITE strsqlx
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

          if (rs(20) & "") = "" OR (RS("A1")& "") ="" or sread <> "" then

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
   IF wspTIPO = "AI" THEN
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","SALVAGUARDAS","si")
   ELSE
    Celda  = Celda & ofv.GenCelda("","TT","","","","","","CONTROLES","si")
   END IF

    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","ut onclick='document." & Form1N & ".submit();' " & wsstyloF,"","10%","","","",sImageOF  ,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","","","","","&nbsp;" & wspDATMdsc,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")








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

   sInput = ofv.gencomboFX("frmasgx",rs(14).name  ,rs(14),scomK,scomD," ")

     IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 )  THEN
     sAccion = "asp/actualizar.asp?fname=" & rs(0)

     wkproccodX = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO&CONTROLTIPO=" & SCONTROLTIPO
     wkproccodX = wkproccodX & "&RIESGO=" & sRIESGO  & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster

      VUELTAX = "../" & REPLACE(wkproccodX,"&","_")
      sAccion = sAccion & "&tabla=ORACTIVRIESGO&volver=" & vueltaX
      else
       saccion=""

      end if
             ' sAccion = sAccion & "&Opcion=SI" & "&CONTROLTIPO=" & SCONTROLTIPO
       Form14  = ofv.genform("frmasgx","AAA STYLE='display:none;' ",sAccion,"","post","",sInput,"si")
    else
       sinput=rs(14)
    end if

    Celda  = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","","",sInput,"si")
    FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


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

	'IF NIVEL >= 3 OR SREQINDICADOR="SI" THEN
   IF SREQINDICADOR="SI" THEN
     sInput2 = "SI " '& NIVEL &"--" & SREQINDICADOR
     strsql="update oractivriesgo set requiereindicador='SI' WHERE OBJID=" & RS(0)
     call ofv.crearconsultaEx(strsql,cn,1,parametros)

   else
      sinput2="NO" '& NIVEL &"--" & SREQINDICADOR
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



    IF (RS(14) & "") = "SI" THEN

        Celda  = ""
        Celda  = Celda & ofv.GenCelda("","uti colspan=4 ","","","","","","OBSERVACIONES","si")
        FilaN = filaN & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

        nCol = 100

        if  len(wdescri) = 0 then wdescri = "Sin observaciones."

        if  not  isnull(rs(15)) then
            wdescri = rs(15)
        else
         wdescri=""
        end if

        t3b = "<font class=at><textarea class=dt "
        t3b = t3b & "name=" & rs(15).name & " cols=" & nCol & " rows=3 "

        t3b = t3b & " onblur=" & sForm & ".submit() "
        t3b = t3b & " onkeypress=" & chr(34)
        t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
        t3b = t3b & " extenso'); return false;}" & chr(34) & " >"

        xdatos =  wdescri
        t3b    = t3b & xdatos & "</textarea></font>"
        Celda  = ofv.GenCelda("","UT colspan=4","center","100%","","","",t3b,"si")
        FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


        Form2N     = "ORCIRCUITONV" & rs(0)
        if sread="" then
         sAccion    = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD & "&RECALC=NO" & "&CONTROLTIPO=1"
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
	             if (rs(20) & "") = "" OR (RS("A2")& "") ="" or sread <> "" then

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
        IF wspTIPO = "AI" THEN
           Celda  = Celda & ofv.GenCelda("","TT style='background-color:#336699;' ","","","","","","SALVAGUARDAS NUEVAS","si")
        ELSE
           Celda  = Celda & ofv.GenCelda("","TT style='background-color:#336699;' ","","","","","","CONTROLES NUEVOS" ,"si")
        END IF

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
            sAccion = sAccion & "&tabla=ORACTIVRIESGO&volver=" & vuelta
         ELSE
            sAccion = ""
         END IF



         FilaN   = ofv.genform(sForm,"",sAccion,"","post","",Filan,"si")





       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF



    Tabla = "<BR><CENTER>" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","80%","","0","","5","2","","",Filan,"si") & "</CENTER>"
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

