<%@LCID = 11274%>
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sopcion  = request.querystring("opcion") & ""
  speriodo = request.querystring("periodo") & ""
  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FORMN"))

  'Encabezados
  '''''Formulario = ofv.Encabezados("descripcion","INICIO","FINAL","ESTADO","")

  Formulario = ""


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORPERIODOS"

  if sopcion = "ADD" then


        strsql = "Insert into ORPERIODOS "
        strsql = strsql & " (PERIODO,INICIO,FINAL,ESTADO,APLIC,"
        strsql = strsql & " RGINSTANCIA,RGVALEAUTOEV,RGVALEACTINF,RGAMBAUTOEV,RGAMBACTINF"
        strsql = strsql & " )"
        strsql = strsql & "  Values "
        strsql = strsql &  "('  NUEVO PERIODO','0','0','NO','" & WKAPLIC & "',"
        strsql = strsql &  "'DEFAULT','S','S','000','001'"
        strsql = strsql &  ")"
        xxx=SqlExLog("INSERT",strsql,CN,"ORPERIODOS","","","PAR")
 '' response.write strsql & "<br>"
       '' call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if










  Strsql = "Select * From " & WKSETTABLA & " WHERE APLIC = '" & WKAPLIC & "' Order By  INICIO DESC,objid desc, FINAL, estado "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir AND ((RS(4) & "") = "NO") then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
     else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if



    'Descripcion
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","5%","","","","NOMBRE","si")
    sInput = rs(1) & ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#336699;' ","","","","","",sInput,"si")




     SELECT CASE (RS(4) & "")
            CASE "NO"
                 sInput5 = "<FONT COLOR='#990000' >NO INICIADO</FONT>"
                 sAccion = "ORPERIODOS.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & rs(0)
                 sAccion = sAccion & "&OPCION=SI&ACCION=OK"
                ' sInput2 = "<INPUT TYPE=BUTTON ONCLICK='document." & sForm & "OK" & ".submit();' VALUE='INICIAR'  STYLE='background-color:#990000;color:#ffffff;' >"
                ' sInput2 = ofv.genform(sForm & "OK","",sAccion,"","post","",sInput2,"si")

                 sAccion = "ORPERIODOSDET.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & rs(0)
                 sInput2 = "<INPUT TYPE=BUTTON ONCLICK='document." & sForm & "DET" & ".submit();' VALUE='PROPIEDADES'  STYLE='background-color:#F7AA22;color:#ffffff;' >"
                 sInput2 = ofv.genform(sForm & "DET","",sAccion,"","post","",sInput2,"si")
                 sInput3 = "&nbsp;"

            CASE "OK"
                 sInput5 = "<FONT COLOR='#009900' >INICIADO</FONT>"
                 sAccion = "ORPERIODOSFIN.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & rs(0)
                 sInput2 = "<INPUT TYPE=BUTTON ONCLICK='document." & sForm & "CL" & ".submit();' VALUE='FINALIZAR'  STYLE='background-color:#009900;color:#ffffff;' >"
                 sInput2 = ofv.genform(sForm & "CL","",sAccion,"","post","",sInput2,"si")
                 sInput3 = "&nbsp;"

            CASE "CL"
                 sInput5 = "<FONT COLOR='#000099' >FINALIZADO</FONT>"
                 sInput2 = "&nbsp;"
                 sInput3 = "&nbsp;"

     END SELECT



    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","5%","","","","MAPA ACTIVOS","si")
    IF RS(6) <> "DEFAULT" AND (RS(10) & "") = "1" THEN
        sInput = "------"
    ELSE

       sInput = rs(10)

      Strsql = "Select A.* "
      Strsql = Strsql & " From RGAMBIENTES A"
      Strsql = Strsql & " where A.CODAMBIENTE = '" & rs(10) & "' and A.TIPOAMBIENTE = 'ACTIVOS'  "
 '''  response.write strsql & "<br>" & chr(13)
      set rsxx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxx.eof then
         sInput = rs(10) & "&nbsp;-&nbsp;" & RSxx(3)
      END IF
    END IF


    Celda = Celda & ofv.gencelda("","UTI","","","","","",sInput,"si")



    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","5%","","","","ACT. IT.","si")
    IF (rs(8) & "") = "S" THEN
       sInput = "<FONT COLOR='#008800' >SI</FONT>"
    ELSE
       sInput = "<FONT COLOR='#880000' >NO</FONT>"
    END IF

    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")




    'fecha
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","5%","","","","INICIO","si")
    if clng(rs(2)) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(rs(2)) + clng(cdate("01-01-1900")) )
    end if

    sInput = sInput1
    Celda = Celda & ofv.gencelda("","VTI ","","","","","",sInput,"si")






    'ESTADO
  '  Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","5%%","","","","ESTADO","si")

    Celda = Celda & ofv.gencelda("","VT style='BORDER-COLOR:#888888;BORDER-WIDTH:1;BORDER-STYLE:SOLID;' ","CENTER","","","","",sInput5,"si")

    Celda = Celda & ofv.GenCelda("","UTI  ","","","","","",sInput3,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")



      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")

    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","INSTANCIA","si")
    IF (rs(6) & "") = "MAPESP" THEN
       sInput = "MAPEO POR PERIODO"
    ELSE
       sInput = "MAPA DE PROCESO PRINCIPAL"
    END IF

    Celda = Celda & ofv.gencelda("","UTI","","","","","",sInput,"si")


    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","MAPA PROCESOS","si")
    IF RS(6) <> "DEFAULT" AND (RS(9) & "") = "0" THEN
        sInput = "------"
    ELSE
       sInput = rs(9)

      Strsql = "Select A.* "
      Strsql = Strsql & " From RGAMBIENTES A"
      Strsql = Strsql & " where A.CODAMBIENTE = '" & rs(9) & "' and A.TIPOAMBIENTE = 'PROCESO'  "
 '''  response.write strsql & "<br>" & chr(13)
      set rsxx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxx.eof then
         sInput = rs(9) & "&nbsp;-&nbsp;" & RSxx(3)
      END IF



    END IF

    Celda = Celda & ofv.gencelda("","UTI","","","","","",sInput,"si")




    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","AUTOEV.","si")
    IF (rs(7) & "") = "S" THEN
       sInput = "<FONT COLOR='#008800' >SI</FONT>"
    ELSE
       sInput = "<FONT COLOR='#880000' >NO</FONT>"
    END IF


    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")


    'fecha
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","FINAL","si")
    if clng(rs(3)) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(rs(3)) + clng(cdate("01-01-1900")) )
    end if

    sInput = sInput1
    Celda = Celda & ofv.gencelda("","VTI ","","","","","",sInput,"si")



    Celda = Celda & ofv.gencelda("","UT  ","CENTER","","","","",sInput2,"si")

    Celda = Celda & ofv.GenCelda("","UTI  ","","","","","","&nbsp;","si")


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","UT COLSPAN=12 style='BORDER-TOP-COLOR:#888888;BORDER-TOP-WIDTH:1;BORDER-TOP-STYLE:SOLID;BORDER-BOTTOM-COLOR:#888888;BORDER-BOTTOM-WIDTH:1;BORDER-BOTTOM-STYLE:SOLID;BACKGROUND-COLOR:#F7EEDD;' ","","","","","","&nbsp;","si")


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Formulario = Formulario & Fila

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
     sAccion = "ORPERIODOS.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=ADD"
     Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select PERIODO From " & WKSETTABLA & " WHERE APLIC = '" & WKAPLIC & "' Order By INICIO,FINAL "    ' orden pagina
  str2 = "Select PERIODO From " & WKSETTABLA & " WHERE APLIC = '" & WKAPLIC & "' Order By PERIODO "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>