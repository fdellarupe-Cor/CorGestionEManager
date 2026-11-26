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





      WKPROCACCDESC = "ACTIVO INFORMATICO"


      sAccion = "RGACTCOMPONENTES.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
      sAccion = sAccion & "&Opcion=SI"


 

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='" & WKSPTITLE & "'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"





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

                  case "AIS"

                        sInput = "ACTIVO INFORMATICO COMPUESTO"
                        sInput2 = "ACTIVO"
                        wspTIPOCLC = "#9900CC"

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



  wkproccod = "../RGACTCOMPONENTES.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster

  VUELTA = REPLACE(wkproccod,"&","_")



  ' Celda = ""
  ' Celda = Celda & ofv.GenCelda("","uti style='background-color:#993333;color:#ffffff;' ","","","","","",sInput2 & ":&nbsp;" & WKPROCACCDESC,"si")
  ' Fila  = ofv.GenRow("","","","","30","","",Celda,"si")
  ' TablaH2 = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;background-color:#aaaaaa;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

    TablaH2 = "" 


''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA RIESGOS ASOCIADOS - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN


        strsql = "Insert into OBJLNK "
        strsql = strsql & " (codobjeto, tipobjeto, iden1, iden2, codobjetoL,"
        strsql = strsql & " tipobjetoL, Lurl, LurlDesc, LTipo,IDEN3L,"
        strsql = strsql & " IDEN4,IDEN4L,POSTXT) "
        strsql = strsql & " Values "
        strsql = strsql & " ('" & sACTIVIDAD & "','" & sTIPOBJETO & "','0','0','0',"
        strsql = strsql & "'0','','','LRGAISD','0',"
        strsql = strsql & " '0','0','0'"
        strsql = strsql & ")"
              '    response.write "10 insert " & strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


 END IF


 IF SOPCION = "DEL" THEN

        strsql = "Delete From OBJLNK Where OBJID = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)


 END IF

''''''''''''''''''''''''''''''''''''''''''''''''''

   Fila = ""

   wspcolor1  = "#990000"
   wspcolor2  = "#009900"
   wspcolor3  = "#D7CF77"
   wspcolor4  = "#EEEEEE"
   sImageDAP  = "<font face='corflaticon' size=3 color='#FFFFFF'>3</font>"







   Strsql = "Select * From OBJLNK where codobjeto = 0" & sACTIVIDAD & " AND tipobjeto = '" & sTIPOBJETO & "' AND LTIPO = 'LRGAISD' "

' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     SNOHAYDSC = "NO EXISTEN COMPONENTES ASOCIADOS AL ACTIVO"


     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","",SNOHAYDSC,"si")

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

    Celda  = Celda & ofv.GenCelda("","uti colspan=2  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","","","","","","COMPONENTE","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","CENTER","5%","","","","DEPENDENCIA","si")

    FilaN  = filaN & ofv.GenRow("","aaa bgcolor='#bbccdf' ","","","25","","",Celda,"si")

    Formulario = FilaN  
    FilaN  = ""

    do until rs.eof

       sForm = "FRMCOMPO" & RS(0)

       wspcolor  = "#336699"
       sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

       wktitdelete = "Confirma la eliminacion deL COMPONENTE asociadO???."

       sHref     = wkproccod
       sLink     = ofv.GenLink("","nt",sHref,"","","",sRIESGODESC,"si")

       Celdak = ""
       Celda  = ""


       wkcoddelete = "S"

       if wkcoddelete = "S"  then
          sImageEL = sImageELS
          sformobj = "FRMRGA" & rs(0)

          sAccion = "RGACTCOMPONENTES.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
          sAccion = sAccion & "&Opcion=DEL&CODRIESGO=" & RS(0)

          formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
          sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "


       else
          wktitdelete = "No se puede eliminar"
          sImageEL = sImageELN
          formobj    = ""
          sfunobj    = ""
       end if


       Celda     = ""
       Celda = Celda & ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2%","","","",sImageEL,"si")


     Strsql = "Select B.CODOBJETO, B.DESCRI From RGACTIVOPRTY A INNER JOIN OBJETOS B ON (A.CODOBJETO = B.CODOBJETO) "
     Strsql = Strsql & " ORDER BY B.DESCRI "
     sInput = ofv.GenerarCombo(sForm,RS(6).NAME,RS(6).value,Strsql,cn,0,1,"")

       Celda  = Celda & ofv.GenCelda("","utI style='border-BOTTOM-style:dotted;border-BOTTOM-width:1;border-BOTTOM-COLOR:#bbbbbb;' ","","70%","","","",sInput,"si")


     Strsql = "Select * From RGACTGRDDEPENDENCIA "
     Strsql = Strsql & " ORDER BY VALOR "
     sInput = ofv.GenerarCombo(sForm,RS(15).NAME,RS(15).value,Strsql,cn,0,2,"")

       Celda  = Celda & ofv.GenCelda("","utI style='border-BOTTOM-style:dotted;border-BOTTOM-width:1;border-BOTTOM-COLOR:#bbbbbb;' ","","30%","","","",sInput,"si")



       FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")


        sAccion = "asp/actualizar.asp?fname=" & rs(0)
       sAccion = sAccion & "&tabla=OBJLNK&volver=" & vuelta
       Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Filan,"si") & formobj

       Filan = ""

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

      rs.movenext
    loop



    Tabla = "<BR>" & ofv.GenTabla("","","","99%","","0","","1","2","","",Formulario,"si")

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

