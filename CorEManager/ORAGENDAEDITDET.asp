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

  svision       = cstr(request.querystring("vision")&"")
  sOPCION       = request.querystring("OPCION")
  sACTIVRIESGO  = request.querystring("ACTIVRIESGO")
  sCODAGENDA    = request.querystring("CODAGENDA")

    ofv.ObtenerAtributos Session("Form"),""
    bAgregar = ofv.AAgregar
    bModificar = ofv.AModificar
    bSuprimir = ofv.AEliminar
    bConsultar = ofv.AConsultar

    set cn = ofv.conectar(ofv.strconn2)

 ' response.write  svision
  if svision="AGNALL" or svision="PORMATIMPACTO" THEN
   SREAD=" READONLY"
  ELSE
   SREAD=""
  END IF




  if sOPCION = "SI" then

        strsql = "Insert into ORAGENDADET "
        strsql = strsql & " (CODTAREA,CODAGENDA,DESCRIPCION,ESTADO, RESPONSABLE,"
        strsql = strsql & " FECINICIO, FECFINAL, DURACION, SECUEN) "
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & GTkey(CLNG(DATE())) & "','" & SCODAGENDA & "','NUEVA TAREA',"
        strsql = strsql &  "'PENDIENTE','0','0','0','1 HS.','0'"
        strsql = strsql & ")"
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if





  SREAD=cstr(SREAD)
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ACTIVRIESGO=" & sACTIVRIESGO & "_VISION=" & SVISION & "_CODAGENDA=" & SCODAGENDA

  sHTML = ofv.FormHeader(session("FormN"))

  Celda = ""


   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","SEGUIMIENTO PLAN DE MITIGACION","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")





    DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))
    dFechaHOY = GTFechadorSh(DHOY)

'''''''''''''''''''''''''''''''''''''''





    xTrColor = "#FFFFFF"

    wkstylo1 = " style='color:#888888;' "
    wkstylo2 = " style='background-color:#336699;color:#ffffff;' "

    wkstyloD = " style='border-style:SOLID;border-width:1;border-color:#888888;' "
    wkstyloT = " style='border-style:SOLID;border-width:1;border-color:#000000;background-color:#F7EEDD;color:#000000;' "

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0110;</font>"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  sImageresp    = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >&#0128;</font>"

  sImagePROC    = "<font face='corflaticon' size=2 color='#888888' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >2</font>"

  sImagePROY    = "<font face='corflaticon' size=2 color='#6699cc' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >&#0204;</font>"

  sImageFNL     = "<font face='corflaticon' size=3 color='#009900' style='border-style:OUTSET;border-width:1;background-color:#F7eedd;' >&#097;</font>"

  'Encabezados
  Formulario = ""


       Celda = ""

       SInput = "TAREA"
       Celda = Celda & ofv.GenCelda("","UTI COLSPAN=2 " & wkstyloT,"","15%","","","",SInput,"si")

       SInput = "DESCRIPCION"
       Celda = Celda & ofv.GenCelda("","UTI " & wkstyloT,"","15%","","","",SInput,"si")

       SInput = "ESTADO"
       Celda = Celda & ofv.GenCelda("","UTI " & wkstyloT,"","15%","","","",SInput,"si")

       SInput = "RESPONSABLE"
       Celda = Celda & ofv.GenCelda("","UTI " & wkstyloT,"","15%","","","",SInput,"si")

       SInput = "DURACION"
       Celda = Celda & ofv.GenCelda("","UTI " & wkstyloT,"","15%","","","",SInput,"si")

       SInput = "INICIO"
       Celda = Celda & ofv.GenCelda("","UTI " & wkstyloT,"","15%","","","",SInput,"si")

       SInput = "FINAL"
       Celda = Celda & ofv.GenCelda("","UTI " & wkstyloT,"","15%","","","",SInput,"si")

    Formulario = Formulario & ofv.GenRow("","","","","25","","",Celda,"si")


	STRSQL = "select estado from oragenda where codagenda = '" & sCODAGENDA & "'"
	set rsEstado = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  WKSETTABLA = "ORAGENDADET"



  Strsql = "Select * From " & WKSETTABLA & " WHERE CODAGENDA = '" & sCODAGENDA & "' ORDER BY SECUEN "
  'Strsql = "select AD.*, a.estado from " & WKSETTABLA & " AD inner join ORAGENDA A ON AD.CODAGENDA = A.CODAGENDA where AD.codagenda = '" & sCODAGENDA & "' ORDER BY AD.SECUEN "

  'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

  posI = 0




   DO UNTIL RS.EOF

    sForm = "form" & posi

    sAccion = ""
    if bModificar then
      sAccion = "ORAGENDAEX.ASP?VISION=UAG&VISIONMAIN=" & SVISION & "&ACTIVRIESGO=" & RS(0)
    end if


   'SECUEN

       Celda = ""


    'Eliminar

    if bSuprimir and sread="" and ((rsEstado("estado").value <> "5" and rsEstado("estado").value <> "0")) then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

       SInput = ofv.GenerarInput(rs(9).name & sread,rs(9),"text",sForm,"5","dtN","")
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","15%","","","",SInput,"si")


      nCol = 60

          wdescri = ""
      if  not  isnull(rs(3)) then
          wdescri = rs(3)
      end if

      t3b = "<font class=at><textarea class=dt "
      t3b = t3b & "name=" & rs(3).name & " cols=" & nCol & " rows=3 "
      if  bModificar AND SREAD="" then
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
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","15%","","","",SInput,"si")



   'ESTADO
       scomK = "PENDIENTE" & chr(9) & "FINALIZADO"
       scomD = scomK
       if sread="" then
        sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4)  ,scomK,scomD," ")
       else
       sInput = rs(4)
       end if

       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","15%","","","",SInput,"si")

   'RESPONSABLE
       sinput="&nbsp;"
       Strsql = "Select ECO.ECO, Eco.descri From Eco ORDER BY Eco.descri "
       if sread = "" then
        sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,Strsql,cn,0,1,"sb")
       else
        Strsql = "Select  Eco.descri From Eco where eco=0" & rs(5)
        set rseco = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rseco.eof then sinput=rseco(0)
        call ofv.cerrarconsulta(rseco)

       end if
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","15%","","","",SInput,"si")

   'DURACION
       SInput = ofv.GenerarInput(rs(8).name & sread ,rs(8) ,"text",sForm,"15","dt","")
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","15%","","","",SInput,"si")

   'FEC.INICIO
      sInputD = ofv.GenerarInput(rs(6).name,rs(6).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(6)) = 0 then
      dFechan = "-------"
    else

      dFechan = GTFechadorSh(rs(6))



    end if

    sInput = dFechan & "&nbsp;" & sInputD

     if sread="" then
      Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;' " & GTCSETCALR(sForm,rs(6).name,GTCSMT,GTCFCH,GTCFCHDAT,"",0,2,0,0,0)  & wkstylo1,"","","","","",sInput,"si")
     else
      Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;' "   & wkstylo1,"","","","","",sInput,"si")
     end if

   'FEC.FINAL
      sInputD = ofv.GenerarInput(rs(7).name,rs(7).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(7)) = 0 then
      dFechan = "-------"
    else

      dFechan = GTFechadorSh(rs(7))


    end if

     sInput = dFechan & "&nbsp;" & sInputD

     if sread="" then
      Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;' " & GTCSETCALR(sForm,rs(7).name,GTCSMT,GTCFCH,GTCFCHDAT,"",0,2,0,0,0)  & wkstylo1,"","","","","",sInput,"si")
     else
     Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;' "   & wkstylo1,"","","","","",sInput,"si")
     end if


     FILA = ofv.GenRow("","AAA BGCOLOR='" & xTrColor & "' ","","","25","","",Celda,"si")

    ' response.write rsEstado("estado")
	if rsEstado("estado").value = "5" or rsEstado("estado").value = "0" then
		sAccion = ""
	else
	    sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
    	sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
	end if

    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",FILA,"si")






       if xTrColor = "#FFFFFF" then
          xTrColor = "#DDDDDD"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    posi = posi + 1

       Formu1 = ""
       Formu2 = ""

   RS.MOVENEXT
  LOOP



 end if
  sHTML = sHTML & "<CENTER><BR>"
  sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#555555;'  ","","90% ","","0","","0","0","","",Formulario,"si")
  sHTML = sHTML & "</CENTER>" & Formu1 & Formu2





  'Botonera
  Agregar = ""
  Buscar  = ""




  if bAgregar and sread="" then
    sAccion = "ORAGENDAEDITDET.asp?OPCION=SI" & "&ACTIVRIESGO=" & sACTIVRIESGO & "&VISION=" & SVISION & "&CODAGENDA=" & SCODAGENDA
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  sAccion = "ORAGENDAEDIT.asp?CantRegMover=" & CantRegMoverx & "&vision=" & svision & "&ACTIVRIESGO=" & sACTIVRIESGO
  Celda = ofv.BotonVolver(sAccion)

	'response.Write(rs)







	if rsEstado("estado").value = "5" or rsEstado("estado").value = "0" then
		sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no","",Buscar,celda)

	else
		sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)
	end if

	call ofv.cerrarconsulta(rsEstado)

  call ofv.cerrarconsulta(rs)


  sHTML = sHTML & GTCCALENDAR("M")





  call ofv.cerrarconn(cn)


    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '''  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"

''    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



Function FechaHAB(wkfecha)


      swd = datepart("w",cdate(wkfecha + clng(cdate("01-01-1900"))))

      IF SWD > 6 THEN
         wkfecha = wkfecha + 2
      END IF

      IF swd = 1 THEN
         wkfecha = wkfecha + 1
      END IF


      FechaHAB  = wkfecha

End Function




%>