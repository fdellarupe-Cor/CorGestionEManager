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
 scodcontin    = request.querystring("codcontin")
 stipocirc    = request.querystring("tipocirc")
 sSHWALL    =  request.querystring("SHWALL")

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



   wkproccod = "PLCIRCUITO.ASP?ACTIVIDAD=" & sACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster 
   Session("ASPSTRPL1") = wkproccod


   wkproccod = "PLCIRCUITO.ASP?ACTIVIDAD=" & sACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster 
   wkproccod = wkproccod & "&CODCONTIN=" & scodcontin & "&TIPOCIRC=" & stipocirc
   Session("ASPSTRPL2") = wkproccod


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

 IF STIPOCIRC <> "" THEN

    WKSPSUBIR = Session("ASPSTRPL1")

 ELSE

    WKSPSUBIR = Session("ASPSTR")

 END IF

 wkproccod = "<a class=nt  href='" & WKSPSUBIR & "' ><b>"
 wkproccod = wkproccod & sImage
 wkproccod = wkproccod & "</b></a>"


           wlnkCOD = 0
           wlnkTIP = "NO"


           Strsql = "SELECT A.RELLINK From ORRSKCIRCUITO A "
           strsql = strsql & " where A.codobjeto = 0" & scodmaster & " AND "
           strsql = strsql & " A.APLIC = '" & wkaplic & "' AND A.PERIODO = 0" & session("ORPERIODO")
 '''  response.write strsql & "<br>" & chr(13)
           set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if  not rsX.eof then

               Strsql = "SELECT CODOBJETOL,TIPOBJETOL From OBJLNK where OBJID = '" & rsX(0) & "'  "
               set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
               if not rsN.eof then
                  wlnkCOD = rsN(0)
                  wlnkTIP = rsN(1)
               end if
               call ofv.cerrarconsulta(rsN)


           END IF
           call ofv.cerrarconsulta(rsX)



 AGREGAR = ""
 wkproccodN = ""

  strsql = "select a.* from OBJACCION a "
  strsql = strsql & " WHERE a.CODACCION = '" &  SACTIVIDAD & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     sCODPROCED = RS(3)
     WKTIPOACC  = "M" & RS(29)

     




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





      WKPROCDESC = rs(10) & "." & sAccionM






  else

      WKPROCDESC = "no se encuentra " &  SACTIVIDAD


  END IF


  IF STIPOCIRC <> "" THEN

     sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:1;' >5</font>"

     IF STIPOCIRC = "PLCIRCA" THEN

       sAccion = Session("ASPSTRPL2")
       sAccion = sAccion & "&Opcion=SI"


       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Agregar circuito alternativo.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"


     ELSE

       sAccion = Session("ASPSTRPL2")
       sAccion = sAccion & "&Opcion=SI"


       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Agregar circuito correctivo.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"


     END IF 



  ELSE

       wkproccodN = ""



     IF sSHWALL = "S" THEN

     sImage = "<font face='corflaticon' size=2 color='#009900' style='border-style:OUTSET;border-width:1;' >0</font>"

       sAccion = Session("ASPSTRPL1")
       sAccion = sAccion & "&SHWALL=N"


       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Ocultar casos de contingencias sin circuito.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"


     ELSE

     sImage = "<font face='corflaticon' size=2 color='#009900' style='border-style:OUTSET;border-width:1;' >1</font>"

       sAccion = Session("ASPSTRPL1")
       sAccion = sAccion & "&SHWALL=S"


       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Mostrar todos los casos de contingencias.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"


     END IF 



  END IF

       wkproccod2 = ""
    WKDIVSECUEN = GENPROCSECUEN()




    Celda = ""
    
   Celda = Celda & ofv.GenCelda("","uti style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","10%","","","","ACTIVIDAD:","si")
   Celda = Celda & ofv.GenCelda("","uti style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","85%","","","",WKPROCDESC,"si")  
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","2%","","","",wkproccodN,"si")  
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","2%","","","",wkproccod2,"si")  
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","5%","","","",wkproccod,"si")  
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   TablaH = ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","100%","0","","-1","1","","",Fila,"si")

''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA RIESGOS ASOCIADOS - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN



    '      wlnkCOD = 0
     '      wlnkTIP = "NO"


   IF STIPOCIRC = "PLCIRCA" THEN

     ' SESSION("WKSECCOD")
     ' SESSION("WKFUNCOD")



     sukey = GTkey(wlnkCOD)

     strsql = "insert into objproced  (codobjeto,tipobjeto,codproced,codeven,frecuencia,tipproced) values (" 
     strsql = strsql & wlnkCOD & ",'" & wlnkTIP & "','" & sukey & "','0','DIARIA','" & STIPOCIRC & "')" 
     call ofv.crearconsultaEx(StrSql,cn,1,volver)



     SCODCIRC = GTkey(SCODMASTER)

     strsql = "insert into PLCIRCALT  ("
     strsql = strsql & "codobjeto,tipobjeto,codcircalt,codcontin,"
     strsql = strsql & "codaccion,codproced,motivo,periodo,"
     strsql = strsql & "ECO,CODPERFIL,MODI,TIPMODI"
     strsql = strsql & ") values (" 
     strsql = strsql & "'" & SCODMASTER & "','" & STIPMASTER & "','" & SCODCIRC & "','" & SCODCONTIN & "',"
     strsql = strsql & "'" & sACTIVIDAD & "','" & sCODPROCED & "','CPP','" & SESSION("ORPERIODO") & "',"
     strsql = strsql & "'0" & SESSION("WKSECCOD") & "','" & SESSION("WKFUNCOD") & "','1','A'"
     strsql = strsql & ")" 
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
 '' response.write strsql & "<br>" & chr(13)

     strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, iden3, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
     strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values ('" & SCODMASTER & "','" & STIPMASTER & "',"
     strsql = strsql & "'" & SCODMASTER & "','" & SCODMASTER & "','" & SCODCIRC & "',"
     strsql = strsql & "'" & wlnkCOD & "','" & wlnkTIP & "','','','LRGPLCA','5','" & sukey & "','0'"
     strsql = strsql & ")"
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

     Strsql = "SELECT * From OBJLNK where iden1 = '" & SCODMASTER & "' and "
     strsql = strsql & " iden2 = '" & SCODMASTER & "' and "
     strsql = strsql & " iden3 = '" & SCODCIRC & "' and "
     strsql = strsql & "LTIPO = 'LRGPLCA' "
     set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsN.eof then
        wlnkid = rsN(0)
     else
        wlnkid = 0
     end if
     call ofv.cerrarconsulta(rsN)

     if clng(wlnkid) > 0 then
        strsql = "update PLCIRCALT set rellink = " & wlnkid
        strsql = strsql & " where  codcircalt = '" & SCODCIRC & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)
     end if




  elseIF STIPOCIRC = "PLCIRCC" THEN



     sukey = GTkey(wlnkCOD)

     strsql = "insert into objproced  (codobjeto,tipobjeto,codproced,codeven,frecuencia,tipproced) values (" 
     strsql = strsql & wlnkCOD & ",'" & wlnkTIP & "','" & sukey & "','0','DIARIA','" & STIPOCIRC & "')" 
     call ofv.crearconsultaEx(StrSql,cn,1,volver)



     SCODCIRC = GTkey(SCODMASTER)

     strsql = "insert into PLCIRCORR ("
     strsql = strsql & "codobjeto,tipobjeto,codcircORR,codcontin,"
     strsql = strsql & "codaccion,codproced,PLAZOMAX,periodo,"
     strsql = strsql & "CODNIVEL,MODI,TIPMODI"
     strsql = strsql & ") values (" 
     strsql = strsql & "'" & SCODMASTER & "','" & STIPMASTER & "','" & SCODCIRC & "','" & SCODCONTIN & "',"
     strsql = strsql & "'" & sACTIVIDAD & "','" & sCODPROCED & "','8','" & SESSION("ORPERIODO") & "',"
     strsql = strsql & "'0','1','A'"
     strsql = strsql & ")" 
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
 '' response.write strsql & "<br>" & chr(13)

     strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, iden3, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
     strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values ('" & SCODMASTER & "','" & STIPMASTER & "',"
     strsql = strsql & "'" & SCODMASTER & "','" & SCODMASTER & "','" & SCODCIRC & "',"
     strsql = strsql & "'" & wlnkCOD & "','" & wlnkTIP & "','','','LRGPLCC','5','" & sukey & "','0'"
     strsql = strsql & ")"
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

     Strsql = "SELECT * From OBJLNK where iden1 = '" & SCODMASTER & "' and "
     strsql = strsql & " iden2 = '" & SCODMASTER & "' and "
     strsql = strsql & " iden3 = '" & SCODCIRC & "' and "
     strsql = strsql & "LTIPO = 'LRGPLCC' "
     set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsN.eof then
        wlnkid = rsN(0)
     else
        wlnkid = 0
     end if
     call ofv.cerrarconsulta(rsN)

     if clng(wlnkid) > 0 then
        strsql = "update PLCIRCORR set rellink = " & wlnkid
        strsql = strsql & " where  codcircORR = '" & SCODCIRC & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)
     end if

  end if 




 END IF


 IF SOPCION = "DEL" THEN

        strsql = "Delete From ORACTIVRIESGO Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "Delete From ORACTIVRCONTROLES Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)




 END IF

''''''''''''''''''''''''''''''''''''''''''''''''''





 if STIPOCIRC = "" THEN





   Fila = ""

   wspcolor1  = "#990000"
   wspcolor2  = "#009900"
   wspcolor3  = "#D7CF77"
   wspcolor4  = "#EEEEEE"
   sImageDAP  = "<font face='corflaticon' size=3 color='#FFFFFF'>3</font>"

  Strsql = ""
  Strsql = Strsql & " SELECT  b.codcontin,c.codcontin,a.* "
  Strsql = Strsql & "   FROM PLTIPOCONT A "
  Strsql = Strsql & "   LEFT JOIN ("
  Strsql = Strsql & "   SELECT DISTINCT codcontin, PERIODO FROM PLCIRCALT Where codaccion = " & SACTIVIDAD & " and  "
  Strsql = Strsql & "   PERIODO = 0" & session("ORPERIODO") & " "
  Strsql = Strsql & "   ) B "
  Strsql = Strsql & "   ON (a.tipocontin = b.codcontin) "
  Strsql = Strsql & "   LEFT JOIN ("
  Strsql = Strsql & "   SELECT DISTINCT codcontin, PERIODO FROM PLCIRCORR Where codaccion = " & SACTIVIDAD & " and "
  Strsql = Strsql & "   PERIODO = 0" & session("ORPERIODO") & " "
  Strsql = Strsql & "   ) C "
  Strsql = Strsql & "   ON (a.tipocontin = c.codcontin) "
'''''''''  Strsql = Strsql & "  WHERE ( B.PERIODO = 0" & session("ORPERIODO") & " ) OR ( C.PERIODO = 0" & session("ORPERIODO") & " )  "
  Strsql = Strsql & " Order By a.TIPOCONTIN "

'  Strsql = " Select  a.codcontin,b.codcontin,c.* From PLTIPOCONT c "
'  Strsql = Strsql & " left join (select distinct codcontin from plcircalt Where codaccion = " & nacccod & ") a "
'  Strsql = Strsql & " on c.tipocontin = a.codcontin "
'  Strsql = Strsql & " left join (select distinct codcontin from plcircorr Where codaccion = " & nacccod & ") b "
'  Strsql = Strsql & " on c.tipocontin = b.codcontin "
'  Strsql = Strsql & " Order By c.TIPOCONTIN "


'   Strsql = "Select * From ORACTIVRIESGO where PERIODO = 0" & session("ORPERIODO") & " AND ACTIVIDAD = '" & SACTIVIDAD & "' "
'   Strsql = Strsql & " ORDER BY SECUEN "

''  response.write strsql & "<br>" & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN CASOS DE CONTINGENCIA DEFINIDOS","si")
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
    Celda  = Celda & ofv.GenCelda("","uti colspan=3  style='background-color:#888888;color:#ffffff;' ","","","","","","RIESGOS ASOCIADOS","si")  

    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#888888;color:#ffffff;' ","CENTER","5%","","","","BRUTO&nbsp;&nbsp;&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#888888;color:#ffffff;' ","CENTER","5%","","","","RESIDUAL","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#888888;color:#ffffff;' ","CENTER","5%","","","","PLAN&nbsp;&nbsp;&nbsp;&nbsp;","si")
  ''  FilaN  = filaN & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

  wkspmastvino = "N"


  do until rs.eof


      wkspsenvino = "N"

      Celda = ofv.GenCelda("","UT","","1%","","","",".","si")


    'Descripcion

    sInput = rs(3)
    Celda = Celda & ofv.GenCelda("","Ut","","90%","","","",sInput,"si")

    


     'Circuito Alternativo

      wkproccod = Session("ASPSTRPL1")
      wkproccod = wkproccod & "&CODCONTIN=" & rs(2) & "&TIPOCIRC=PLCIRCA"


      if (rs(0) & "") = "" then
          xclas  = "btbn"
          sInput = "<font class='btic'  >&#0115;</font>"

      else
          wkspsenvino = "S"
          sInput = "<font class='btic'  >&#097;</font>"
          xclas  = "btbs"
      end if


      sHref = wkproccod

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = ""
      stitle = "Circuito Alternativo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


     'Circuito Correctivo

      wkproccod = Session("ASPSTRPL1")
      wkproccod = wkproccod & "&CODCONTIN=" & rs(2) & "&TIPOCIRC=PLCIRCC"

      if (rs(1) & "") = "" then
          xclas  = "btbn"
          sInput = "<font class='btic'  >&#0115;</font>"

      else
          wkspsenvino = "S"
          sInput = "<font class='btic'  >&#097;</font>"
          xclas  = "btbs"
      end if


      sHref = wkproccod

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = ""
      stitle = "Circuito Correctivo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      wkspclass = "RNV" 

      if wkspsenvino = "S" then 
         wkspclass = "RSV" 
         wkspmastvino = "S"
      end if

    Fila= ofv.GenRow("",wkspclass & " bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")

    if xTrColor = "#FFFFFF" then
       xTrColor = "#F0F0F0" 
    ELSE
       xTrColor = "#FFFFFF" 
    END IF



    Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")


    rs.movenext
  loop


  TABLA = ofv.GenTabla("","","","99%","","0","","0","0","","",Formulario,"si")


  if wkspmastvino = "N" or sSHWALL = "S" then

  else


    sHTML = sHTML & "<style>"
    sHTML = sHTML & " TR.RNV {DISPLAY:'NONE';}"
    sHTML = sHTML & "</style>"


  end if


  end if





 ELSE

           WKSPCONTINGENCIA = "SD"
           WKSPCONTINASOCDSC = "SD"
           WKSPCONTINASOCCOD = 0


           Strsql = "SELECT A.* From PLTIPOCONT A "
           strsql = strsql & " where A.tipocontin = 0" & SCODCONTIN
 '''  response.write strsql & "<br>" & chr(13)
           set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if  not rsX.eof then

               WKSPCONTINGENCIA  = RSX(1)
               WKSPCONTINASOCCOD = RSX(2)

           END IF
           call ofv.cerrarconsulta(rsX)


           Celda = ""
    

           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","CONTINGENCIA:","si") 
           Celda = Celda & ofv.GenCelda("","utI  ","","80%","","","",WKSPCONTINGENCIA,"si") 

           Fila  = ofv.GenRow("","","","","25","","",Celda,"si")





       
      wspcolor1  = "#990000"
      wspcolor2  = "#009900"
      wspcolor3  = "#D7CF77"
      wspcolor4  = "#EEEEEE"
      sImageDAP  = "<font face='corflaticon' size=3 color='#FFFFFF'>3</font>"

       wkcoddelete = "N"
       wktitdelete = " "
       sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
       sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"

       xTrColor = "#FFFFFF"


   VUELTA = "../" & REPLACE(Session("ASPSTRPL2"),"&","_")

   IF STIPOCIRC = "PLCIRCA" THEN


       TABLAA = ofv.GenTabla("","AAA style='background-color:#ddddff;border-style:SOLID;border-width:1;Border-color:#000000;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

      Fila = ""

     Strsql = ""
     Strsql = Strsql & "   SELECT * FROM PLCIRCALT Where codaccion = " & SACTIVIDAD & " and  "
     Strsql = Strsql & "   PERIODO = 0" & session("ORPERIODO") & " AND CODCONTIN = 0" & SCODCONTIN

''  response.write strsql & "<br>" & chr(13)
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs.eof then


        Celda = ""
        Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN CIRCUITOS ALTERNATIVOS DEFINIDOS<BR> PARA ESTE CASO DE CONTINGENCIA.","si")
        Fila = ofv.GenRow("","","","","","","",Celda,"si")
        Tablab = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

        Tablab = "<CENTER><BR><BR><BR><BR>" & Tablab & "</CENTER>"

     else


       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","ut style='background-color:#888888;color:#ffffff;' ","","","","","","#","si")  
       Celda  = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","","","","","MOTIVO","si")  

       Celda  = Celda & ofv.GenCelda("","UtI  style='background-color:#888888;color:#ffffff;' ","","","","","","AREA/SECTOR","si")
       Celda  = Celda & ofv.GenCelda("","UtI  style='background-color:#888888;color:#ffffff;' ","","","","","","PERFIL","si")
       Celda  = Celda & ofv.GenCelda("","UtI  style='background-color:#888888;color:#ffffff;' ","","","","","","DETALLE","si")
       Formulario = ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

       POSI = 0

       do until rs.eof

           POSI = POSI + 1


          sform = "FRMSEC" & POSI




         Celda  = ""

         Celda  = Celda & ofv.GenCelda("","ut style='background-color:#993333;color:#ffffff;' ","CENTER","1%","","","",POSI,"si")  

        scomK = "CPP" & chr(9) & "CPO" & chr(9) & "CPC"
        scomD = "CONTINGENCIA PROPIA" & chr(9) & "CONTINGENCIA PRESENTADA EN AREA ORIGINANTE" & chr(9) & "CONTINGENCIA PRESENTADA EN AREA CONTINUADORA"
        sInput = ofv.gencomboFX(SForm,rs(18).name,RS(18),scomK,scomD," ")

         Celda  = Celda & ofv.GenCelda("","uti  ","","","","","",sInput,"si")  


          Strsql1 = "Select * From Eco Order By descri"
          sInput = ofv.GenerarCombo(sform,RS(4).NAME,RS(4),Strsql1,cn,0,1,"")
         Celda  = Celda & ofv.GenCelda("","UtI  ","","","","","",sInput,"si")


         Strsql1 = "SELECT TIPOPERFIL.tipoperfil,TIPOPERFIL.descri "
          Strsql1 = Strsql1 & "FROM ECO, ECOPERFIL, TIPOPERFIL "
          Strsql1 = Strsql1 & " where  "
          Strsql1 = Strsql1 & " ECO.eco = ECOPERFIL.codsector and "
          Strsql1 = Strsql1 & " ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
          Strsql1 = Strsql1 & " and Eco.eco = 0" & RS(4)

          sInput = ofv.GenerarCombo(sform,RS(16).NAME,RS(16),Strsql1,cn,0,1,"")
         Celda  = Celda & ofv.GenCelda("","UtI  ","","","","","",sInput,"si")


         Form1N     = "PLCIRCUITODET" & rs(0)

         wkproccodCC = "PLCIRCUITODET.ASP?ACTIVIDAD=" & sACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster 
         wkproccodCC = wkproccodCC & "&CODCONTIN=" & scodcontin & "&TIPOCIRC=" & stipocirc & "&codCIRC=" & RS(3)
 
         Form1      = ofv.genform(Form1N,"AAA STYLE='display:none;' ",wkproccodCC,"","post","","","si")


         wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;"
         wsstyloF   = wsstyloF & "background-color:#f7aa22;' "

         sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >&#0232;</font>"
         Celda  = Celda & ofv.GenCelda("","Ut title='ACTIVIDADES DEL CIRCUITO.' onclick='document." & Form1N & ".submit();' " & wsstyloF,"CENTER","2","","","",sImageOF,"si")


         Fila= ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","30","","",Celda,"si")

         if xTrColor = "#FFFFFF" then
            xTrColor = "#F0F0F0" 
         ELSE
            xTrColor = "#FFFFFF" 
         END IF


         sAccion = "asp/actualizar.asp?fname=" & rs(0)
         sAccion = sAccion & "&tabla=PLCIRCALT&volver=" & vuelta

         Formulario = Formulario & ofv.GenForm(sform,"",sAccion,"","post","",Fila,"si") & Form1


         rs.movenext
       loop

       TABLAB = ofv.GenTabla("","","","99%","","0","","0","0","","",Formulario,"si")

     end if


   ELSE

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' CIRCUITO CORRECTIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Strsql = "Select DESCRI From Eco WHERE ECO = 0" & WKSPCONTINASOCCOD
          set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rs.eof then
             WKSPCONTINASOCDSC = RS(0)
          END IF


           Celda = ""
           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","ASOCIACION:","si") 
           Celda = Celda & ofv.GenCelda("","utI  ","","80%","","","",WKSPCONTINASOCDSC,"si") 

           Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

       TABLAA = ofv.GenTabla("","AAA style='background-color:#ddddff;border-style:SOLID;border-width:1;Border-color:#000000;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

      Fila = ""

     Strsql = ""
     Strsql = Strsql & "   SELECT * FROM PLCIRCORR Where codaccion = " & SACTIVIDAD & " and  "
     Strsql = Strsql & "   PERIODO = 0" & session("ORPERIODO") & " AND CODCONTIN = 0" & SCODCONTIN

''  response.write strsql & "<br>" & chr(13)
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs.eof then


        Celda = ""
        Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN CIRCUITOS CORRECTIVOS DEFINIDOS<BR> PARA ESTE CASO DE CONTINGENCIA.","si")
        Fila = ofv.GenRow("","","","","","","",Celda,"si")
        Tablab = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

        Tablab = "<CENTER><BR><BR><BR><BR>" & Tablab & "</CENTER>"

     else


       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","ut style='background-color:#888888;color:#ffffff;' ","","","","","","#","si")  
       Celda  = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","","","","","NIVEL","si")  
       Celda  = Celda & ofv.GenCelda("","UtI  style='background-color:#888888;color:#ffffff;' ","","","","","","DETALLE","si")
       Formulario = ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

       POSI = 0

       do until rs.eof

           POSI = POSI + 1


          sform = "FRMSEC" & POSI




         Celda  = ""

         Celda  = Celda & ofv.GenCelda("","ut style='background-color:#993333;color:#ffffff;' ","CENTER","1%","","","",POSI,"si")  

 

          Strsql1 = "Select * From PLNIVELES WHERE CODASOC = 0" & WKSPCONTINASOCCOD & " Order By SECUEN"
          sInput = ofv.GenerarCombo(sform,RS(4).NAME,RS(4),Strsql1,cn,0,1,"")
         Celda  = Celda & ofv.GenCelda("","UtI  ","","","","","",sInput,"si")



         Form1N     = "PLCIRCUITODET" & rs(0)

         wkproccodCC = "PLCIRCUITODET.ASP?ACTIVIDAD=" & sACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster 
         wkproccodCC = wkproccodCC & "&CODCONTIN=" & scodcontin & "&TIPOCIRC=" & stipocirc & "&codCIRC=" & RS(3)
 
         Form1      = ofv.genform(Form1N,"AAA STYLE='display:none;' ",wkproccodCC,"","post","","","si")


         wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;"
         wsstyloF   = wsstyloF & "background-color:#f7aa22;' "

         sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >&#0232;</font>"
         Celda  = Celda & ofv.GenCelda("","Ut title='ACTIVIDADES DEL CIRCUITO.' onclick='document." & Form1N & ".submit();' " & wsstyloF,"CENTER","2","","","",sImageOF,"si")


         Fila= ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","30","","",Celda,"si")

         if xTrColor = "#FFFFFF" then
            xTrColor = "#F0F0F0" 
         ELSE
            xTrColor = "#FFFFFF" 
         END IF

         sAccion = "asp/actualizar.asp?fname=" & rs(0)
         sAccion = sAccion & "&tabla=PLCIRCORR&volver=" & vuelta

         Formulario = Formulario & ofv.GenForm(sform,"",sAccion,"","post","",Fila,"si") & Form1


         rs.movenext
       loop

       TABLAB = ofv.GenTabla("","","","99%","","0","","0","0","","",Formulario,"si")

     end if

   END IF 



   TABLA = TABLAA & TABLAB


 END IF








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
  Celda = Celda & TABLA
  Celda = Celda & "</div>"

  sHTML = sHTML & Celda

 
'''''''''''''''''''''''''''''''''''''''''''''

  sHTML = sHTML & WKDIVSECUEN & Chr(13)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,"no","")

  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML


 FUNCTION GENPROCSECUEN()

    GENPROCSECUEN = ""

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

       FOR WKI=WKITOPE TO 7 

           Celda = ""
           Celda2 = ""

               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")

               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")

           FOR WKIC=1 TO WKINCOL

               WKISTYLO = ""

               IF WKIC = WKINCOL THEN

                  WKISTYLO = "border-LEFT-style:SOLID;border-LEFT-width:1;Border-LEFT-color:#F7AA22;"
                  WKISTYLO = WKISTYLO & "border-BOTTOM-style:SOLID;border-BOTTOM-width:1;Border-BOTTOM-color:#F7AA22;"

               END IF

               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;" & WKISTYLO & "' ","","","","","","&nbsp;","si")

               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")


           NEXT  

           Celda = Celda & ofv.GenCelda("","aaa colspan=4 ROWSPAN=2 bgcolor='#ffffff' style='font-family:tahoma;font-size:7pt;color:#888888;border-style:solid;border-width:1;border-color:#993333;'  ","","","","","","&nbsp;" & WRS(WKI),"si")


           FOR WKIC=WKINCOL TO WKICCOL

               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")

               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")


           NEXT  

           FilaW = FilaW & ofv.GenRow("","","","","5","","",Celda,"si")
           FilaW = FilaW & ofv.GenRow("","","","","5","","",Celda2,"si")

           WKINCOL = WKINCOL + 1

           WKISTYLO = ""

           IF WKINCOL < (WKICCOL + 1) THEN

              WKISTYLO = "border-LEFT-style:SOLID;border-LEFT-width:1;Border-LEFT-color:#F7AA22;"

           END IF


           Celda = ""
           Celda = Celda & ofv.GenCelda("","aaa colspan=" & ((WKINCOL * 2) ) & " style='font-family:tahoma;font-size:4pt;' ","","","","","","&nbsp;","si")
           Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
           Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;" & WKISTYLO & "' ","","","","","","&nbsp;","si")
           Celda = Celda & ofv.GenCelda("","aaa colspan=" & (((WKICCOL - WKINCOL) * 2) + 6) & " style='font-family:tahoma;font-size:4pt;' ","","","","","","&nbsp;","si")
           FilaW = FilaW & ofv.GenRow("","","","","5","","",Celda,"si")




       NEXT


       IF WKITOPE < 8 THEN

           Celda = ""
           Celda = Celda & ofv.GenCelda("","aaa colspan=" & ((WKICCOL * 2) + 8) & " style='font-family:tahoma;font-size:4pt;' ","","","","","","&nbsp;","si")
           FilaW = ofv.GenRow("","","","","5","","",Celda,"si") & FilaW

          FilaW  = ofv.GenTabla("","aaa ","","100%","","0","","-1","-1","","",FilaW,"si")

          sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0120;</font>"

          wkproccod3 = "<a class=nt onclick=""document.all.DX5.style.display='none';"" href='javascript:void(0);' "
          wkproccod3 = wkproccod3 & " ><b>"
          wkproccod3 = wkproccod3 & sImage
          wkproccod3 = wkproccod3 & "</b></a>"


          Celda = ""
          Celda = Celda & ofv.GenCelda("","uti ","","98%","","","","ENCADENAMIENTO DEL PROCESO","si")
          Celda = Celda & ofv.GenCelda("","ut ","","2%","","","",wkproccod3,"si")
          FilaW2 = ofv.GenRow("","","","","","","",Celda,"si")
          FilaW  = ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' ","","100%","","0","","0","1","","",FilaW2,"si") & FilaW

          FilaW  = "<CENTER>" & FilaW & "</CENTER>"



          Celda = "" & Chr(13)


          Celda = Celda & "<div id='DX5' style='z-index:1;margin-top:1;margin-left:1;position:absolute;left:400px;top:140px;display:none;"
          Celda = Celda & "width:400;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
          Celda = Celda & "background-color:#F9F9F9;"
          Celda = Celda & "' >" & Chr(13)



          Celda = Celda & FilaW & Chr(13)


          Celda = Celda & "</div>" & Chr(13)


          sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:1;' >4</font>"

          wkproccod2 = "<a class=nt onclick=""document.all.DX5.style.display='block';"" href='javascript:void(0);' "
          wkproccod2 = wkproccod2 & " title='ENCADENAMIENTO DEL PROCESO.' >"
          wkproccod2 = wkproccod2 & sImage
          wkproccod2 = wkproccod2 & "</a>"


          GENPROCSECUEN = Celda

       END IF



    END IF


 END FUNCTION

%>

