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
 scodcirc    = request.querystring("codcirc")

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







   wkproccod = "PLCIRCUITODET.ASP?ACTIVIDAD=" & sACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster 
   wkproccod = wkproccod & "&CODCONTIN=" & scodcontin & "&TIPOCIRC=" & stipocirc & "&codCIRC=" & scodcirc
   Session("ASPSTRPL3") = wkproccod


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



    WKSPSUBIR = Session("ASPSTRPL2")



 wkproccod = "<a class=nt  href='" & WKSPSUBIR & "' ><b>"
 wkproccod = wkproccod & sImage
 wkproccod = wkproccod & "</b></a>"



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

       sAccion = Session("ASPSTRPL3")
       sAccion = sAccion & "&Opcion=SI"


       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Agregar actividad del circuito alternativo.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"


     ELSE

       sAccion = Session("ASPSTRPL3")
       sAccion = sAccion & "&Opcion=SI"


       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Agregar actividad del circuito correctivo.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"


     END IF 



  ELSE

       wkproccodN = ""

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
           Celda = Celda & ofv.GenCelda("","utI colspan=3 ","","80%","","","",WKSPCONTINGENCIA,"si") 

           Fila  = ofv.GenRow("","","","","25","","",Celda,"si")





     IF STIPOCIRC = "PLCIRCA" THEN

        Strsql = ""
        Strsql = Strsql & "   SELECT * FROM PLCIRCALT Where CODCIRCALT = '" & SCODCIRC & "' "

''  response.write strsql & "<br>" & chr(13)
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then


           wkspmoticod = rs(18) & ""
           wkspecocod = "0" & rs(4)
           wkspfuncod = rs(16) & ""


           select case wkspmoticod

                  case "CPP"

                        wkspmotiDSC = "CONTINGENCIA PROPIA"

                  case "CPO"

                        wkspmotiDSC = "CONTINGENCIA PRESENTADA EN AREA ORIGINANTE"

                  case "CPC"

                        wkspmotiDSC = "CONTINGENCIA PRESENTADA EN AREA CONTINUADORA"

                  case ELSE

                        wkspmotiDSC = "DESCONOCIDA"

           end select





          Strsql = "Select DESCRI From Eco WHERE ECO = " & wkspecocod
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then
             wkspecoDSC = RSX(0)
          ELSE
             wkspECODSC = "SIN ASOCIAR"

          END IF

          Strsql = "Select DESCRI From TIPOPERFIL WHERE TIPOPERFIL.tipoperfil = '" & wkspfuncod & "' "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then
             wkspFUNDSC = RSX(0)
          ELSE
             wkspFUNDSC = "SIN ASOCIAR"

          END IF


           
          Strsql = "SELECT * From OBJLNK where OBJID = 0" & RS(8)
          set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsN.eof then
             wlnkPROCED = rsN(12)
          else
             wlnkPROCED = 0
          end if
          call ofv.cerrarconsulta(rsN)




                     '  wkspmotiDSC = wkspmotiDSC & "-" & SCODCIRC & "-" & RS(8) & "-" & wlnkPROCED & "-" & wlnkCOD & "-" & wlnkTIP

        end if


     ELSE


          Strsql = "Select DESCRI From Eco WHERE ECO = 0" & WKSPCONTINASOCCOD
          set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rs.eof then
             WKSPCONTINASOCDSC = RS(0)
          END IF


           Celda = ""
           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","ASOCIACION:","si") 
           Celda = Celda & ofv.GenCelda("","utI colspan=3 ","","80%","","","",WKSPCONTINASOCDSC,"si") 

           Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")



        Strsql = ""
        Strsql = Strsql & "   SELECT * FROM PLCIRCORR Where CODCIRCORR = '" & SCODCIRC & "' "

''  response.write strsql & "<br>" & chr(13)
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then



           wkspNIVcod = "0" & rs(4)

          Strsql = "Select DESCRI From PLNIVELES WHERE CODNIVEL = 0" & wkspNIVcod
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then
             wkspNIVDSC = RSX(0)
          ELSE
             wkspNIVDSC = "SIN ASOCIAR"

          END IF


           Celda = ""
           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","NIVEL:","si") 
           Celda = Celda & ofv.GenCelda("","utI colspan=3 ","","80%","","","",wkspNIVDSC,"si") 

           Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")



        end if


     END IF 






    ' TABLAA = ofv.GenTabla("","AAA style='background-color:#ddddff;border-style:SOLID;border-width:1;Border-color:#000000;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")





''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA ACTIVIDADES ASOCIADOS - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN



    '      wlnkCOD = 0
     '      wlnkTIP = "NO"




     ' SESSION("WKSECCOD")
     ' SESSION("WKFUNCOD")




                GETFECH()
  
                WKFECHA = XLFECDAT - clng(cdate("01-01-1900"))
                WKHORA  = XLHORDAT

                WKFORDEN = 10


                strsql = "SELECT MAX(SECUEN) FROM OBJACCION "
                strsql = strsql & " WHERE CODOBJETO = " & wlnkCOD & " AND TIPOBJETO = '" & wlnkTIP & "' "
                strsql = strsql & " AND CODPROCED = '" & wlnkPROCED & "' "
                set rsN2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if NOT rsN2.eof then

                   WKFORDEN = rsN2(0) + 10

                end if


                sevento = 0
                sukey = GTkey(wlnkCOD)

                strsql = "insert into objaccion  ("
                strsql = strsql & "codobjeto,tipobjeto,codproced,codaccion,"
                strsql = strsql & "secuen,tipproced,descriacc,MODI,"
                strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA,"
                strsql = strsql & "CRITICO"
                strsql = strsql & ") values (" 
                strsql = strsql & wlnkCOD & ",'" & wlnkTIP & "','" & wlnkPROCED & "','" & sukey & "',"
                strsql = strsql & "'" & WKFORDEN & "','" & STIPOCIRC & "','NUEVA ACTIVIDAD','1',"
                strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "',"
                strsql = strsql & "'SI'" 
                strsql = strsql & ")" 
                call ofv.crearconsultaEx(StrSql,cn,1,volver)






 END IF


 IF SOPCION = "DEL" THEN

    SACCIONCONT = request.querystring("ACCIONCONT")

        strsql = "Delete From OBJACCION Where CODACCION = '" & SACCIONCONT & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

      '  strsql = "Delete From ORACTIVRCONTROLES Where CODRIESGO = '" & SCODRIESGO & "' "
      '  call ofv.crearconsultaEx(StrSql,cn,1,volver)




 END IF

''''''''''''''''''''''''''''''''''''''''''''''''''












       
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

   IF STIPOCIRC = "PLCIRCA" THEN



           Celda = ""
    

           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","MOTIVO:","si") 
           Celda = Celda & ofv.GenCelda("","utI COLSPAN=3 ","","80%","","","",wkspmotiDSC,"si") 

           Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

           Celda = ""

           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","AREA/SECTOR:","si") 
           Celda = Celda & ofv.GenCelda("","utI  ","","30%","","","",wkspECODSC,"si") 
           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","PERFIL/FUNCION:","si") 
           Celda = Celda & ofv.GenCelda("","utI  ","","30%","","","",wkspFUNDSC,"si") 

           Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")









   ELSE

           Celda = ""
    

           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","MOTIVO:","si") 
           Celda = Celda & ofv.GenCelda("","utI COLSPAN=3 ","","80%","","","",wkspmotiDSC,"si") 

       '    Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

           Celda = ""

           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","AREA/SECTOR:","si") 
           Celda = Celda & ofv.GenCelda("","utI  ","","30%","","","",wkspmotiDSC,"si") 
           Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","20%","","","","PERFIL/FUNCION:","si") 
           Celda = Celda & ofv.GenCelda("","utI  ","","30%","","","",wkspmotiDSC,"si") 

         '  Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


   END IF 


           TABLAA = ofv.GenTabla("","AAA style='background-color:#ddddff;border-style:SOLID;border-width:1;Border-color:#000000;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

      Fila = ""

   VUELTA = "../" & REPLACE(Session("ASPSTRPL3"),"&","_")
   VUELTAC = REPLACE(Session("ASPSTRPL3"),"&","_")


           strsql = "select a.objid,a.objid,a.* from OBJACCION a "
           strsql = strsql & " WHERE a.codproced = '" &  wlnkPROCED & "' "
           strsql = strsql & " and (a.modi = 0 or (a.modi <> 0 and a.tipmodi <> 'E')) "
           strsql = strsql & " Order By a.secuen"



 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO HAY ACTIVIDADES PARA EL CIRCUITO ACTUAL","si")
     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     TablaN = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

     TablaB = "<CENTER><BR><BR>" & TablaN & "</CENTER>"

  else





    xTrColor = "#FFFFFF"

    wkcant = 1
    wkleft = 5
    wktop  = 5


 
       scomK = "SI" & chr(9) & "NO"
        scomD = "SI" & chr(9) & "NO"

       CeldaK  = ""
       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","uti colspan=5 ","","","","","","ACTIVIDADES","si")  
       Celda  = Celda & ofv.GenCelda("","UtI colspan=5 ","","","","","","OPCIONES","si")


       FilaN = filaN & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")



       wkcoddelete = "N"
       wktitdelete = " "
       sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
       sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"


    do until rs.eof

       wkcoddelete = "N"
       wktitdelete = " "

 

       Form2N     = "ORCRITICIDAD" & rs(0)




       wkproccod = "javascript:void(0);"

       sHref     = wkproccod
       sImage    = "<font face='corflaticon' size=1 color='#FFFFFF'>&#060;</font>"
       sLink     = ofv.GenLink("","nt",sHref,"","","",sImage,"si")


       wspcolor  = "#F7AA22"
       sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

   IF STIPOCIRC = "PLCIRCA" THEN

         Strsql = "SELECT tipoaccionalt,descri From PLTipoaccalt Order By descri"

   ELSE

         Strsql = "SELECT tipoaccioncorr,descri From PLTIPOACCCORR Order By descri"

   END IF  



          sInput = ofv.GenerarCombo(Form2N,rs(6).NAME,rs(6).value,Strsql,cn,0,1,"")

          wkproccod2 = ofv.GenerarInput(rs(12).NAME,rs(12),"text",Form2N,"3","dt","") & "."
          wkproccod = sInput


          wktitdelete = "Confirma la eliminacion de la actividad???."
          wkcoddelete = "S"



       ANCHOK = "99%"
       Celdak = ""
       Celda  = ""

       if wkcoddelete = "S" then
          sImageEL = sImageELS
          sformobj = "FRMOBJ" & rs(0)
          sAccion    = Session("ASPSTRPL3") & "&OPCION=DEL&ACCIONCONT=" & RS("CODACCION")
          formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
          sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" " 


       else
          sImageEL = sImageELN
          formobj    = ""
          sfunobj    = ""
       end if




       Celda     = ""
       Celda = Celda & ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageEL,"si")





       Celda  = Celda & ofv.GenCelda("","ut","CENTER","2","","","",sImageP,"si")  
       Celda  = Celda & ofv.GenCelda("","UT bgcolor=" & wspcolor & " ","center","2","","","",sLink,"si") 
       Celda  = Celda & ofv.GenCelda("","UT ","","2","","","",wkproccod2,"si") 
       Celda = Celda & ofv.GenCelda("","UtI","",ANCHOK,"","","",wkproccod,"si")





  'Texto de referencia

      sID = "TextoEsquema.asp?"
      sID = sID & "Accion=" & RS(0) & "&Codproced=" & wlnkPROCED & "&Codobjeto=" & wlnkCOD
      sID = sID & "&Tipobjeto=" & wlnkTIP & "&volver=" & vuelta

      if rs(13) <> "" or rs(18) <> "" or rs(19) <> "" then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = ""

      stitle = "Actualizar texto de referencia."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

   IF STIPOCIRC = "PLCIRCA" THEN


  'Asociacion de Complementos

      sID = "EdGENLnkCOMP.asp?"
      sID = sID & "ID=" & RS(0) & "&TABLA=OBJACCION&Codobjeto=" & wlnkCOD
      sID = sID & "&Tipobjeto=" & wlnkTIP & "&volver=" & vueltaC

      if cstr(rs(8) & "") <> "0" then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0163;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0163;</font>"
      end if

       sHref = sID
      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = ""

      stitle = "ASOCIAR INFORMACION DE APOYO."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


  'TIPO DISPARADOR

       Celda = Celda & ofv.GenCelda("","UtI","","","","","","&nbsp;","si")





       

  'Continuador


      'Sector Continuador

      sID = "EdProcedCON.asp?"
      sID = sID & "Accion=" & RS(0) & "&Codproced=" & wlnkPROCED & "&Codobjeto=" & wlnkCOD
      sID = sID & "&Tipobjeto=" & wlnkTIP & "&volver=" & VUELTA

      if  cstr(rs(10)) <> "0"  then
          xclas  = "btbs"
          sInput = "<font class='btic' >G</font>"
      else
          xclas  = "btbn"
          sInput = "<font class='btic' >G</font>"
      end if
       sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = ""
      stitle = "Establecer Sector Continuador."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      'perfil Continuador

      if  cstr(rs(10)) <> "0"  then
          sID = "EdProcedPER.asp?"
          sID = sID & "Accion=" & RS(0) & "&Codproced=" & wlnkPROCED & "&Codobjeto=" & wlnkCOD
          sID = sID & "&Tipobjeto=" & wlnkTIP & "&volver=" & VUELTA

          sper = rs(11)

          if isnull(sper) then
             xclas  = "btbn"
             sInput = "<font class='btic' >&#0128;</font>"
          elseif sper = " " then
             xclas  = "btbn"
             sInput = "<font class='btic' >&#0128;</font>"
          else
             xclas  = "btbs"
             sInput = "<font class='btic' >&#0128;</font>"
          end if

          sHref = sID

      else
          sHref  = "javascript:void(0);"
          xclas  = "btbn"
          sInput = "<font class='btic' >&#0128;</font>"
      end if


      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = ""
      stitle = "Establecer Perfil Continuador."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


   
   END IF

 
   IF STIPOCIRC = "PLCIRCC" THEN

      Celda = Celda & ofv.GenCelda("","UtI","","","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","UtI","","","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","UtI","","","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","UtI","","","","","","&nbsp;","si")


   END IF




'''''''''




 




       FilaX = ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

       sAccion = "asp/actualizar.asp?fname=" & rs(0)
       sAccion = sAccion & "&tabla=OBJACCION&volver=" & vuelta
       FilaN      = filaN & ofv.genform(Form2N,"AAA STYLE='display:none;' ",sAccion,"","post","",FilaX,"si") & Form1 & formobj


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF




      rs.movenext
    loop





       TablaB = ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")




   END IF





   TABLA = TABLAA & TABLAB











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

