<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 response.expires = 0

 sACTIVIDAD = request.querystring("codmaster")
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

  strsql = "select a.* from OBJETOS a "
  strsql = strsql & " WHERE a.CODOBJETO = '" &  SACTIVIDAD & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     sCODPROCED = RS(0)
   '  WKTIPOACC  = "M" & RS(29)

     








      WKPROCDESC = rs(3)


      sAccion = "ORCIRCUITOINFO.asp?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster 
      sAccion = sAccion & "&Opcion=SI"
  '    Agregar = ofv.BotonAgregar(sAccion)

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:1;' >5</font>"

       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Agregar riesgo asociado.'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"



  else

      WKPROCDESC = "no se encuentra " &  SACTIVIDAD


  END IF




    Celda = ""
    
   Celda = Celda & ofv.GenCelda("","uti style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","10%","","","","ACTIVO:","si")
   Celda = Celda & ofv.GenCelda("","uti style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","85%","","","",WKPROCDESC,"si")  
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","2%","","","",wkproccodN,"si")  
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","","5%","","","",wkproccod,"si")  
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   TablaH = ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","100%","0","","-1","1","","",Fila,"si")

''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA RIESGOS ASOCIADOS - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN



        strsql = "Insert into ORACTIVRIESGO "
        strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
        strsql = strsql & " SECUEN, CODRIESGO, ITEMRIESGO, PROBABILIDAD, IMPACTO, "
        strsql = strsql & " RIESGOBRUTO, RIESGORESIDUAL, VULNERABILIDAD,REQUIEREPLAN) "
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & session("ORPERIODO") & "','" & sACTIVIDAD & "','" & sCODPROCED & "',"
        strsql = strsql &  "'" & SCODMASTER & "','" & STIPMASTER & "',"
        strsql = strsql & "'10','" & GTkey(SCODMASTER) & "',"
        strsql = strsql & "'0','0','0','0','0','0','NO'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)


 END IF


 IF SOPCION = "DEL" THEN

        strsql = "Delete From ORACTIVRIESGO Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "Delete From ORACTIVRCONTROLES Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)




 END IF

''''''''''''''''''''''''''''''''''''''''''''''''''

   Fila = ""

   wspcolor1  = "#990000"
   wspcolor2  = "#009900"
   wspcolor3  = "#D7CF77"
   wspcolor4  = "#EEEEEE"
   sImageDAP  = "<font face='corflaticon' size=3 color='#FFFFFF'>3</font>"

   Strsql = "Select * From ORACTIVRIESGO where PERIODO = 0" & session("ORPERIODO") & " AND ACTIVIDAD = '" & SACTIVIDAD & "' "
   Strsql = Strsql & " ORDER BY SECUEN "

 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN RIESGOS ASOCIADOS AL ACTIVO","si")
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
    Celda  = Celda & ofv.GenCelda("","uti colspan=3  style='background-color:#336699;color:#ffffff;' ","","","","","","AMENAZAS ASOCIADOS","si")  

    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#336699;color:#ffffff;' ","CENTER","5%","","","","BRUTO&nbsp;&nbsp;&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#336699;color:#ffffff;' ","CENTER","5%","","","","RESIDUAL","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='background-color:#336699;color:#ffffff;' ","CENTER","5%","","","","PLAN&nbsp;&nbsp;&nbsp;&nbsp;","si")
    FilaN  = filaN & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

    do until rs.eof


       wspcolor  = "#336699"
       sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

       wkproccod = "ORCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD
       wkproccod = wkproccod & "&RIESGO=" & RS(7) & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster


       WKCODRIESGO = "NO"
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

    IF RS(8) > 0 THEN WKCODRIESGO = "SI"


       sHref     = wkproccod
       sLink     = ofv.GenLink("","nt",sHref,"","","",sRIESGODESC,"si")

       Celdak = ""
       Celda  = ""

       wktitdelete = "Confirma la eliminacion del riesgo asociado???."
       wkcoddelete = "S"

       if wkcoddelete = "S" then
          sImageEL = sImageELS
          sformobj = "FRMRGA" & rs(0)
       '''   sAccion    = "respproceso.asp?codobjeto=" & rs(7) & "&tipobjeto=" & "RGA" & ""

          sAccion = "ORCIRCUITOINFO.asp?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster 
          sAccion = sAccion & "&Opcion=DEL&CODRIESGO=" & RS(7)

          formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
          sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" " 


       else
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

          wspDATA     = "---"            
          wspDATB     = "---" 
          wspDATP     = "---" 

       ELSE




          wspDATA     = rs(11)           
          wspDATB     = rs(12)
          wspDATP     = rs(14)


          if rs(11) = 0  then

             wspcolor0A  = wspcolor4

          elseif rs(11) < 9  then

             wspcolor0A  = wspcolor2

          elseif rs(11) < 18  then

             wspcolor0A  = wspcolor3

          elseif rs(11) > 17  then

             wspcolor0A  = wspcolor1

          end if  


          if rs(12) = 0  then

             wspcolor0B  = wspcolor4

          elseif rs(12) < 9  then

             wspcolor0B  = wspcolor2

          elseif rs(12) < 18  then

             wspcolor0B  = wspcolor3

          elseif rs(12) > 17  then

             wspcolor0B  = wspcolor1

          end if  


          if rs(14) = "SI"  then

             wspcolor0P  = wspcolor2

          else

             wspcolor0P  = wspcolor4

          end if  


      

       END IF




       wsstyloF   = " style='border-style:OUTSET;border-width:1;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0A & ";' "

       Celdak  = Celdak & ofv.GenCelda("","Vt " & wsstyloF,"CENTER","5%","","","",wspDATA,"si")

       wsstyloF   = " style='border-style:OUTSET;border-width:1;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0B & ";' "

       Celdak  = Celdak & ofv.GenCelda("","Vt " & wsstyloF,"CENTER","5%","","","",wspDATB,"si")

       wsstyloF   = " style='border-style:OUTSET;border-width:1;COLOR:#FFFFFF;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";' "

       Celdak  = Celdak & ofv.GenCelda("","Ut " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")


       Celda = Celda & Celdak
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
  Celda = Celda & TABLA
  Celda = Celda & "</div>"

  sHTML = sHTML & Celda

 
'''''''''''''''''''''''''''''''''''''''''''''



  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,"no","")

  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML

%>

