<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  sestado = 1
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar


  vuelta = ""


 ' response.write "modif " & bModificar
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  WKOPCION = request.querystring("PASO")

 '''' session("PASOOBS") = WKOPCION
 ''' RESPONSE.WRITE  session("PASOOBS")

  IF NCODIGO <> "" THEN
     session("NCliente") = NCODIGO
  ELSE
     NCODIGO = session("NCliente")
  END IF

  IF WKOPCION <> "" THEN
     session("PASOOBS") = WKOPCION
  ELSE
     WKOPCION = session("PASOOBS")
  END IF

  IF sTipo <> "" THEN
     session("TCliente") = sTipo
  ELSE
     sTipo = session("TCliente")
  END IF

  vuelta = ofv.vueltaasp2("../") & vuelta
  vuelta2 = ofv.vueltaasp2("")
  sestado = 9

  set cn = ofv.conectar(ofv.strconn2)


 IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN

  Strsql = "SELECT ESTADO FROM ORRSKCIRCUITO  "
  Strsql = Strsql & " Where OBJID = 0" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sestado = clng(rs("estado"))
  call ofv.cerrarconsulta(rs)

 ELSE


  Strsql = "SELECT * FROM OBJprty  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sestado = clng(rs("estado"))
  call ofv.cerrarconsulta(rs)

 END IF



  if CSTR(sestado) > "2" and CSTR(sestado) < "5" then
     if session("Nomodificar") = "S" then bModificar = False
     response.redirect "coobjwfldet.asp?cod=SI&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo

  elseif (CSTR(sestado) = "2" or CSTR(sestado) = "5") and Session("Form2") <> "DESARROLLO" AND WKAPLIC <> "OR" AND WKAPLIC <> "PL" THEN
 '''''''' elseif (CSTR(sestado) = "2" or CSTR(sestado) = "5" or CSTR(sestado) = "6")  THEN
     response.redirect "edobjwflconf.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

  else
  if not (session("segct") = "S") then
     if session("Nomodificar") = "S" then bModificar = False
  end if

  sHTML   = ofv.FormHeader(session("FormN"))



  IF WKOPCION = "OBS" or WKOPCION = "APRO" THEN
     WKOPCION = "APRO"
     WKOPCIONdsc = "OBSERVACION"
  ELSE
     WKOPCION = "AUTO"
     WKOPCIONdsc = "AUTORIZACION"
  END IF


  session("CKTTYPE") = WKOPCION


  Strsql = "SELECT * FROM OBJWFL  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  rs.eof then

      strsql = "Insert into OBJWFL (codobjeto, tipobjeto, ordenest,esqauto,chat,msgchat) "
      strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','1','1','1','SE ENCUENTRA ACTIVO PARA LA AUTOEVALUACION EL PRESENTE CIRCUITO.')"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)

      strsql = "Insert into OBJWFLDET (codobjeto, tipobjeto, orden,USR,COMEN1,CONF,CKTTYPE) "
      strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','1','NO','REFERENTE','SC','" & session("CKTTYPE") & "')"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)

      strsql = "Insert into OBJWFLDET (codobjeto, tipobjeto, orden,USR,COMEN1,CONF,CKTTYPE) "
      strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','2','NO','SUPERVISORREFERENTE','SC','" & session("CKTTYPE") & "')"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)


      strsql = "Insert into OBJWFLDET (codobjeto, tipobjeto, orden,USR,COMEN1,CONF,CKTTYPE) "
      strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','3','NO','DUEÑOPROCESO','SC','" & session("CKTTYPE") & "')"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)



  else
   ''  response.write "v " & rs(6) & "<br>"
      if rs(6)  then response.redirect "edObjWflHst.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

  end if

  call ofv.cerrarconsulta(rs)



  xActiva = 0
  xObliga = 0

 IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN

  xActiva = 1
  xObliga = 1

 ELSE

  Strsql = "SELECT * FROM Defwfl"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then

     xActiva = rs(1)
     xObliga = rs(2)

  end if

  call ofv.cerrarconsulta(rs)

 END IF



    'Estado
     dim Aest(6)
     Aest(0) = 0
     Aest(1) = "Desarrollo"
     Aest(2) = "Para Aprobar"
     Aest(3) = "Aprobada"
     Aest(4) = "Eliminada"

     Strsql = "SELECT * from estados  "
     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do until rsx.eof
        Aest(rsx(0)) = Ucase(rsx(1))
        rsx.movenext
     loop
     call ofv.cerrarconsulta(rsx)


    SCODMASTER   = 0
    SPROCMAIN    = 0
    SPROCMAINCOD = 0

     Strsql = "SELECT CODOBJETO from ORRSKCIRCUITO WHERE OBJID = 0" & nCodigo
     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     IF NOT rsx.eof THEN SCODMASTER = rsx(0)
     call ofv.cerrarconsulta(rsx)



     Strsql = "SELECT * FROM OBJWFL  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    Celda = ofv.gencelda("","ttt colspan=4 aaa ","center","100%","","","","<BIG>CIRCUITO DE " & WKOPCIONdsc & "</BIG>","si")
    Fila  = ofv.GenRow("","","","","30","","",Celda,"si")

    Celda  = ofv.gencelda("","tt","","20%","","","","ESTADO&nbsp;ACTUAL:","si")
    sinput = ofv.gencelda("","vti colspan=3","","","","","",Aest(CLNG(sestado)),"si")
    Celda  = celda & sinput
    Fila   = FILA & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda  = ofv.gencelda("","tt","","20%","","","","SUBPROCESO:","si")
    sinput = ofv.gencelda("","vti colspan=3","","","","","",GENPROCSECUENEX(SCODMASTER),"si")
    Celda  = celda & sinput
    Fila   = FILA & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda  = ofv.gencelda("","tt","","20%","","","","PROCESO:","si")
    sinput = ofv.gencelda("","vti colspan=3","","","","","",SPROCMAIN,"si")
    Celda  = celda & sinput
    Fila   = FILA & ofv.GenRow("","","","","25","","",Celda,"si")




  if not rs.eof then

                   WKFSIGOK = "S"

''''''''''''''''''''''''''''''''''''''''''''''''
        filaS1 = ""

        Strsql = "SELECT * FROM OBJwfldet  "
        Strsql = Strsql & " Where codobjeto=" & nCodigo
        Strsql = Strsql & " And tipobjeto = '" & stipo & "' And CKTTYPE = '" & session("CKTTYPE") & "' order by orden,objid"
        response.write strsql
        set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT RSX1.EOF THEN

           DO WHILE NOT RSX1.EOF

            IF (RSX1(4) & "") = "NO" THEN WKFSIGOK = "N"

            SINPUT = "RESPONSABLE"

            SELECT CASE (RSX1(3) & "")

                   CASE "1"
                     SINPUT = "REFERENTE:"

                   CASE "2"
                     SINPUT = "SUPERVISOR REFERENTE:"

                   CASE "3"
                     SINPUT = "DUEÑO DEL PROCESO:"

            END SELECT


            Celda1 = ofv.gencelda("","UT style='border-style:solid;border-width:1;border-color:#999999;' ","","30%","","","",SINPUT,"si")

            sval = (RSX1(4) & "")

                             scomK = scomK & "NO" & chr(9)
                             scomD = scomD & "SIN ASIGNAR" & chr(9)




                          '   strsql = "Select * From USUARIOS    "
                          '   strsql = strsql & " WHERE ID IN ("
                          '   wkfsep3 = ""
                          '   strsql2 = "Select DISTINCT GRPUSR From OBJSDKDET WHERE codobjeto = 0" & SPROCMAINCOD & "  "
                          '   set rsniv = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
                          '   DO UNTIL RSNIV.EOF
                          '      WKUSRSEL = REPLACE(RSNIV(0) & "","USR-","")
                          '      strsql = strsql & wkfsep3 & "'" & WKUSRSEL & "'"
                          '      wkfsep3 = ","
                          '   RSNIV.MOVENEXT
                          '   LOOP

                          '   strsql = strsql & " )"
                          '   strsql = strsql & " Order By DESCRIPCION"
 ''  response.write strsql & "<br>" & chr(13)
                          '   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                          '   DO UNTIL rsniv.eof
                          '      xxx=getinfouser(rsniv(0))
                          '      xnombre=session("SGU_nombre")
                          '      NSELECT = NSELECT & "<OPTION VALUE='"  & rsniv(0) & "' >"  & xnombre  & "</OPTION>"
                          '      scomK = scomK & rsniv(0) & chr(9)

                          '      scomD = scomD &  xnombre & chr(9)
                          '     RSNIV.MOVENEXT
                          '   LOOP
                          '   call ofv.cerrarconsulta(rsniv)



                             strsql2 = "Select DISTINCT GRPUSR From OBJSDKDET WHERE codobjeto = 0" & SPROCMAINCOD & "  "
                             set rsniv = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
                             DO UNTIL RSNIV.EOF
                              '  WKUSRSEL = REPLACE(RSNIV(0) & "","USR-","")
                              '  strsql = strsql & wkfsep3 & "'" & WKUSRSEL & "'"
                              '  wkfsep3 = ","

                                xxx=getinfouser(REPLACE(RSNIV(0) & "","USR-",""))
                                xnombre=session("SGU_nombre")
                                NSELECT = NSELECT & "<OPTION VALUE='"  & REPLACE(RSNIV(0) & "","USR-","") & "' >"  & xnombre  & "</OPTION>"
                                scomK = scomK & REPLACE(RSNIV(0) & "","USR-","") & chr(9)

                                scomD = scomD &  xnombre & chr(9)



                                RSNIV.MOVENEXT
                             LOOP
                             call ofv.cerrarconsulta(rsniv)



                             NSELECT = NSELECT & "</SELECT> "


                      	     sInput2 = ofv.gencomboFX(sForm & "AUT" & rsX1(0),rsX1(4).NAME,rsX1(4),scomK,scomD,"")


            Celda1 = Celda1 & ofv.gencelda("","UT style='border-style:solid;border-width:1;border-color:#999999;' ","","70%","","","",sInput2,"si")

            XFTROW =  ofv.GenRow("","","","","25","","",Celda1,"si")

            sAccion = "asp/ActualizarEN.asp?ID=" & rsX1(0) & "&Tabla=ObjwflDET&volver=" & vuelta
            sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
            filaS1 = filaS1 & ofv.GenForm(sForm & "AUT" & rsX1(0),"",sAccion,"","post","",XFTROW,"si")


          RSX1.MOVENEXT
          LOOP

        ELSE

           WKFSIGOK = "N"



        END IF



      CeldaT = ofv.GenTabla("","aaa  ","","90% style='border-style:solid;border-width:1;border-color:#999999;' ","","0","0","1","2","","",filaS1,"si")

     Celda = ofv.gencelda("","UT COLSPAN=4 ","CENTER","","","","",CeldaT,"si")
     Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

''''''''''''''''''''''''''''''''''''''''''''''''

   sForm = "wfl"

    'Texto del Mensaje
  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:transparent;} "

     XMESAURL   =  left(request.servervariables("SCRIPT_NAME"),instr(RIGHT(request.servervariables("SCRIPT_NAME"),LEN(request.servervariables("SCRIPT_NAME"))-1),"/"))

     XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL

       Celda = ofv.gencelda("","ttr colspan=4","LEFT","","","","",XMESAURL ,"si")
     '  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")



       Celda = ofv.gencelda("","ttr colspan=4","LEFT","","","","","MENSAJE:","si")
       Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

       if bModificar  and (Cint(sestado) = 1 or Cint(sestado) = 6) then
    '''   if bModificar  and (Cint(sestado) = 1 ) then

          wkfmensaje = rs(7) & ""

          if wkfmensaje = "" then

            IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN

             wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un CIRCUITO"
             wkfmensaje = wkfmensaje & " pendiente de su aprobación."
             wkfmensaje = wkfmensaje & CHR(13) & CHR(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
             wkfmensaje = UCASE(wkfmensaje)

            ELSE

             wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un proyecto"
             wkfmensaje = wkfmensaje & " normativo pendiente de su aprobación."
             wkfmensaje = wkfmensaje & CHR(13) & CHR(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
             wkfmensaje = UCASE(wkfmensaje)

            END IF


             strsql = "update OBJWFL set MSGCHAT = '" & wkfmensaje & "'  "
             Strsql = Strsql & " Where codobjeto=" & nCodigo
             Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if

          nCol = 80

          t3b = "<font class=at><textarea CLASS=DT "
          t3b = t3b & "name=" & rs(7).name & " cols=" & nCol & " rows=6 "
          if bModificar then
             t3b = t3b & " onblur=" & sForm & "3.submit() "
             t3b = t3b & " onkeypress=" & chr(34)
             t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
             t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
          else
             t3b = t3b & " readonly>"
          end if
          t3b = t3b & wkfmensaje & "</textarea></font>"
          sInput = ofv.gencelda("","ut colspan=4","center","","","","",t3b,"si")
          sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
          sInput =  ofv.GenForm(sForm & "3","",sAccion,"","post","",sInput,"si")
       else
          sinput = ofv.gencelda("","ut colspan=4","center","","","","",rs(7),"si")
       end if

       Celda = sinput
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


   end if


       WKFSIG = "S"


    if rs(3) then
       WKFSIG2 = "S"
    ELSE
       WKFSIG2 = "N"
    END IF

    if rs(4) then
       WKFMODAL = "S"
    ELSE
       WKFMODAL = "N"
    END IF

 end if


  call ofv.cerrarconsulta(rs)

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aaa  ","","70% style='border-style:solid;border-width:1;border-color:#999999;' ","","0","0","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"

  FILA = ""

  Celda = ofv.gencelda("","ttt colspan=4","","","","","","&nbsp;estado del circuito","si")
  Formulario2 = ofv.GenRow("","","","","30","","",Celda,"si")

  if WKFSIG = "S" THEN

     wkxhay = "N"

     set cn5 = ofv.conectar(ofv.strconn5)

     StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & session("TCliente") & "' and  OBJCOD = '" & session("NCliente") & "' AND FIN = 'NO'  "
     set rsxx = ofv.crearconsultaEx(StrSql,cn5,1,parametros)
     if not rsxx.eof then
        wkxhay = "S"

     END IF
     call ofv.cerrarconn(cn5)

     wktabpermite = split(session("WKSPTPERMCE") & "","-")
     wktabcircuito = ubound(wktabpermite)
     wkpermisoce = wktabpermite(wktabcircuito)




   '''  IF (Cint(sestado) = 1 or Cint(sestado) = 6 ) AND (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S"  and wkxhay <> "S" then

   'response.write " ncliente:" & session("NCliente") & "-tcliente:" & session("TCliente") & "-sestado:" & sestado & "-bmodificar:" & bmodificar & "-nomodificarce:" & session("Nomodificarce") & "-segce:" & wktabcircuito & "/" & wkpermisoce & "/" & session("WKSPTPERMCE") & "-xhay:" & wkxhay & "<br>"

     IF (Cint(sestado) = 1 ) AND (bModificar   or session("Nomodificarce") <> "S") and wkpermisoce = "S"   and wkxhay <> "S"  then

    '    Strsql = "SELECT * FROM OBJwfldet  "
    '    Strsql = Strsql & " Where codobjeto=" & nCodigo
    '    Strsql = Strsql & " And tipobjeto = '" & stipo & "' And CKTTYPE = '" & session("CKTTYPE") & "' order by orden,objid"
    '    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    '    IF RS.EOF THEN
    '       WKFSIGOK = "N"
    '    ELSE
    '       WKFSIGOK = "S"
    '    END IF


        IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN

           Strsql = "SELECT * FROM ORRSKCONFIG WHERE APLIC = '" & WKAPLIC & "'  "
           set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           IF RS.EOF THEN
              WKFSIGOK = "N2"
           ELSE
              IF WKOPCION = "AUTO" THEN
                 WKFCKT = RS(5)
              ELSE
                 WKFCKT = RS(6)
              END IF
           END IF

        ELSE

           Strsql = "SELECT * FROM TIPOBJETO  "
           Strsql = Strsql & " Where "
           Strsql = Strsql & " tipobjeto = '" & stipo & "' "
           set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           IF RS.EOF THEN
              WKFSIGOK = "N2"
           ELSE
              IF WKOPCION = "AUTO" THEN
                 WKFCKT = RS(19)
              ELSE
                 WKFCKT = RS(20)
              END IF
           END IF

        end if



        IF WKFSIGOK = "S" or (WKFSIGOK = "N" and WKFSIG2 = "N" AND WKOPCION = "AUTO" ) THEN

           Celda = ""
           IF WKFSIGOK = "N" THEN
              sInput2 = "<b>ATENCION!!</b><BR>EN CASO DE INICIAR, EL CIRCUITO CERRARA CON "
              SINPUT2 = SINPUT2 & " <u><b>" & SESSION("SEGUSUARIO") & "</b></u>  COMO UNICO AUTORIZANTE.<BR>"
           ELSE
              sInput2 = "<b>ATENCION!!</b><BR>EL CIRCUITO ESTA EN CONDICIONES DE SER INICIADO POR "
              SINPUT2 = SINPUT2 & " <u><b>" & SESSION("SEGUSUARIO") & "</b></u>  COMO ORIGINANTE.<BR>"
           END IF
           Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT2,"si")

           sInput = ofv.GenerarInput("boton1","INICIAR","submit","edPropiedades2","5","bt","")
           Celda = celda & ofv.gencelda("","uti","center","10%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")

           sAccion = "EDWKFSTART.asp?OPCION=" & WKOPCION & "&MODAL=" & WKFMODAL
           Formulario2 = Formulario2 &  ofv.GenForm("edPropiedades2","",sAccion,"","post","",Fila,"si")

        ELSE

           Celda = ""
           IF WKFSIGOK = "N" THEN
              sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, NO INFORMO LA LISTA DE AUTORIZANTES.<BR>"
           ELSE
              sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, EL TIPO DE DOCUMENTO NO ESTA CONFIGURADO.<BR>"
           END IF
           Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT,"si")

           sInput = "&nbsp;"
           Celda = celda & ofv.gencelda("","ut","center","5%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")

           Formulario2 = Formulario2 &  Fila

        END IF

    ELSE

           Celda = ""

          IF (Cint(sestado) = 1 ) AND (bModificar   or session("Nomodificarce") <> "S") and wkpermisoce = "S" and wkxhay = "S"  then
             sInput = "<FONT COLOR=#005500>EN CURSO.</FONT><BR>"
          else
             sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO DESDE AQUI O NO ESTA AUTORIZADO PARA INICIARLO.<BR>"
          end if
           Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT,"si")

           sInput = "&nbsp;"
           Celda = celda & ofv.gencelda("","ut","center","5%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")

           Formulario2 = Formulario2 &  Fila


    END IF

''     IF (Cint(sestado) = 1 ) AND (bModificar   or session("Nomodificarce") <> "S") and wkpermisoce = "S"   and wkxhay <> "S"  then

 ''   response.write " ncliente:" & session("NCliente") & "-tcliente:" & session("TCliente") & "-sestado:" & sestado & "-bmodificar:" & bmodificar & "-nomodificarce:" & session("Nomodificarce") & "-segce:" & wktabcircuito & "/" & wkpermisoce & "/" & session("WKSPTPERMCE") & "-xhay:" & wkxhay & "<br>"

     IF ((Cint(sestado) = 2 or Cint(sestado) = 5 or Cint(sestado) = 6 ) OR (Cint(sestado) = 1 and wkxhay = "S"  ) ) AND  wkpermisoce = "S"   then

         '  sInput = ofv.GenerarInput("boton1","Consultar historial","submit","edPropiedades3","5","bt","")
         '  sAccion = "CoHistWkf.asp"
         '  sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
         '  Celda = ofv.gencelda("","Vt colspan=2","center","30%","","","",sInput,"si")

           if (bModificar   ) and wkpermisoce = "S" then
              sInput = ofv.GenerarInput("boton1","Cancelar Circuito","submit","edPropiedades3","5","bt style='background-color:#993333;' ","")
              sAccion = "EDCANCELWkf.asp?OPCION=CONF"
              sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
           else
              sInput = "&nbsp;"

           end if


           Celda = ""

           Celda = celda & ofv.gencelda("","Vt colspan=2 ","center","30%","","","",sInput,"si")

           sInput = "&nbsp;"
           Celda = celda & ofv.gencelda("","Vt colspan=2","center","30%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")


           Formulario2 = Formulario2 & Fila


     end if




 END IF

'xxxxxxxxxx



  IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN

     sAccion = Session("ASPSTR") & "&ACCIONEJEC=ABREESTAD"
     Retornar = ofv.BotonVolver(sAccion)


  END IF


  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#999999;' ","","70% ","","0","0","1","0","","",Formulario2,"si")
  sHTML = sHTML & "</center>"

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  sestado = 1
  sHTML = sHTML & "</body></html>"
'' end if
end if

  if (Cint(sestado) = 1 or Cint(sestado) = 6)  then response.write sHTML

 FUNCTION GENPROCSECUENEX(scodmaster)

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

           XCelda = ""
           XCelda2 = ""

       SPROCMAIN    =  WRS(WKITOPE)
       SPROCMAINCOD =  WRS(WKITOPE + 8)

       FOR WKI=WKITOPE TO 7




           XCelda = XCelda & XCelda2 & WRS(WKI)

           XCelda2 = "<b><font color='#00aa00' size=2 >&#0124;</font></b>"



       NEXT


       IF WKITOPE < 8 THEN

          GENPROCSECUENEX = XCelda & XCelda2

       END IF



    END IF


 END FUNCTION

%>