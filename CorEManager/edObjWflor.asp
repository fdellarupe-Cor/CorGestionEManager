
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = ofv.CheckUsuario()
If ok <> "Y" Then
    sestado = 1
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    sHTML = ok
Else
    response.expires = 0

    XXXX = GENLOGEXT("EDIDET")
    APLCHECK()
	CantRegMoverx = cdbl( "0" & request.querystring("CantRegMover"))
    'IF session("APLMDCLI")="BC" THEN

    ofv.ObtenerAtributos Session("Form2"), ""
    bModificar = ofv.AModificar


    vuelta = ""


    '  response.write request.querystring
    nCodigo = request.querystring("Codobjeto")
    stipo = request.querystring("Tipobjeto")

    WKOPCION = request.querystring("PASO")

    '''' session("PASOOBS") = WKOPCION
    ''' RESPONSE.WRITE  session("PASOOBS")

    If nCodigo <> "" Then
        Session("NCliente") = nCodigo
    Else
        nCodigo = Session("NCliente")
    End If

    If WKOPCION <> "" Then
        Session("PASOOBS") = WKOPCION
    Else
        WKOPCION = Session("PASOOBS")
    End If

    If stipo <> "" Then
        Session("TCliente") = stipo
    Else
        stipo = Session("TCliente")
    End If

    'vuelta = ofv.vueltaasp2("../") & vuelta &"_ddd=4"
	vuelta = ofv.vueltaasp("../", cantregmostrar, CantRegMoverx)
    vuelta2 = ofv.vueltaasp2("")
    sestado = 9



    Set cn = ofv.conectar(ofv.strconn2)


    If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

        StrSql = "SELECT ESTADO FROM ORRSKCIRCUITO  "
        StrSql = StrSql & " Where OBJID = 0" & nCodigo
        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not rs.EOF Then sestado = CLng(rs("estado"))
        Call ofv.cerrarconsulta(rs)

    Else


        StrSql = "SELECT * FROM OBJprty  "
        StrSql = StrSql & " Where codobjeto=" & nCodigo
        StrSql = StrSql & " And tipobjeto = '" & stipo & "'"
        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not rs.EOF Then sestado = CLng(rs("estado"))
        Call ofv.cerrarconsulta(rs)

    End If



    If CStr(sestado) > "2" And CStr(sestado) < "5" Then
        If Session("Nomodificar") = "S" Then bModificar = False
        response.redirect "coobjwfldet.asp?cod=SI&codobjeto=" & nCodigo & "&tipobjeto=" & stipo

    ElseIf (CStr(sestado) = "2" Or CStr(sestado) = "5") And Session("Form2") <> "DESARROLLO" And WKAPLIC <> "OR" And WKAPLIC <> "PL" Then
        '''''''' elseif (CSTR(sestado) = "2" or CSTR(sestado) = "5" or CSTR(sestado) = "6")  THEN
        response.redirect "edobjwflconf.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & stipo

    Else
        If Not (Session("segct") = "S") Then
            If Session("Nomodificar") = "S" Then bModificar = False
        End If

        sHTML = ofv.FormHeader(Session("FormN"))
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------


        If WKOPCION = "OBS" Or WKOPCION = "APRO" Then
            WKOPCION = "APRO"
            WKOPCIONdsc = "OBSERVACION"
        Else
            WKOPCION = "AUTO"
            WKOPCIONdsc = "AUTORIZACION"
        End If


        Session("CKTTYPE") = WKOPCION

		IF cint(request.querystring("idUpdate")) <> 0 and request.querystring("estadoactivo") = "N" THEN
			StrSql = "UPDATE OBJWFLDET SET USR = 'NO' WHERE OBJID = 0" & request.querystring("idUpdate")
			Call ofv.crearconsultaEx(StrSql, cn, 1, volver)
		END IF


        StrSql = "SELECT * FROM OBJWFL  "
        StrSql = StrSql & " Where codobjeto=" & nCodigo
        StrSql = StrSql & " And tipobjeto = '" & stipo & "'"
        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If rs.EOF Then

            StrSql = "Insert into OBJWFL (codobjeto, tipobjeto, ordenest,esqauto,chat,msgchat) "
            StrSql = StrSql & " Values (" & nCodigo & ",'" & stipo & "','1','1','1','SE ENCUENTRA ACTIVO PARA LA AUTOEVALUACION EL PRESENTE CIRCUITO.')"
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)

            StrSql = "Insert into OBJWFLDET (codobjeto, tipobjeto, orden,USR,COMEN1,CONF,CKTTYPE,ESTADOPASOWKF) "
            StrSql = StrSql & " Values (" & nCodigo & ",'" & stipo & "','1','NO','REFERENTE','SC','" & Session("CKTTYPE") & "','S')"
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)

            StrSql = "Insert into OBJWFLDET (codobjeto, tipobjeto, orden,USR,COMEN1,CONF,CKTTYPE,ESTADOPASOWKF) "
            StrSql = StrSql & " Values (" & nCodigo & ",'" & stipo & "','2','NO','SUPERVISORREFERENTE','SC','" & Session("CKTTYPE") & "','S')"
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)


            StrSql = "Insert into OBJWFLDET (codobjeto, tipobjeto, orden,USR,COMEN1,CONF,CKTTYPE,ESTADOPASOWKF) "
            StrSql = StrSql & " Values (" & nCodigo & ",'" & stipo & "','3','NO','DUEÑOPROCESO','SC','" & Session("CKTTYPE") & "','S')"
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)



        Else
            ''  response.write "v " & rs(6) & "<br>"
            If rs(6) Then response.redirect "edObjWflHst.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & stipo

        End If

        Call ofv.cerrarconsulta(rs)



        xActiva = 0
        xObliga = 0

        If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

            xActiva = 1
            xObliga = 1

        Else

            StrSql = "SELECT * FROM Defwfl"
            Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

            If Not rs.EOF Then

                xActiva = rs(1)
                xObliga = rs(2)

            End If

            Call ofv.cerrarconsulta(rs)

        End If



        'Estado
        Dim Aest(6)
        Aest(0) = 0
        Aest(1) = "Desarrollo"
        Aest(2) = "Para Aprobar"
        Aest(3) = "Aprobada"
        Aest(4) = "Eliminada"

        StrSql = "SELECT * from estados  "
        Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        Do Until rsx.EOF
            Aest(rsx(0)) = UCase(rsx(1))
            rsx.MOVENEXT
        Loop
        Call ofv.cerrarconsulta(rsx)


        scodmaster = 0
        SPROCMAIN = 0
        SPROCMAINCOD = 0

        StrSql = "SELECT CODOBJETO from ORRSKCIRCUITO WHERE OBJID = 0" & nCodigo
        Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not rsx.EOF Then scodmaster = rsx(0)
        Call ofv.cerrarconsulta(rsx)



        StrSql = "SELECT * FROM OBJWFL  "
        StrSql = StrSql & " Where codobjeto=" & nCodigo
        StrSql = StrSql & " And tipobjeto = '" & stipo & "'"
        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

        'response.write (Chr(13))
        'response.write (StrSql & Chr(13))
        'response.write (Chr(13))


        Celda = ofv.gencelda("", "ttt colspan=4 aaa ", "center", "100%", "", "", "", "<BIG>CIRCUITO DE " & WKOPCIONdsc & "</BIG>", "si")
        Fila = ofv.GenRow("", "", "", "", "30", "", "", Celda, "si")

        Celda = ofv.gencelda("", "tt", "", "20%", "", "", "", "ESTADO&nbsp;ACTUAL:", "si")
        sInput = ofv.gencelda("", "vti colspan=3", "", "", "", "", "", Aest(CLng(sestado)), "si")
        Celda = Celda & sInput
        Fila = Fila & ofv.GenRow("", "", "", "", "25", "", "", Celda, "si")

        Celda = ofv.gencelda("", "tt", "", "20%", "", "", "", "SUBPROCESO:", "si")
        sInput = ofv.gencelda("", "vti colspan=3", "", "", "", "", "", GENPROCSECUENEX(scodmaster), "si")
        Celda = Celda & sInput
        Fila = Fila & ofv.GenRow("", "", "", "", "25", "", "", Celda, "si")

        Celda = ofv.gencelda("", "tt", "", "20%", "", "", "", "PROCESO:", "si")
        sInput = ofv.gencelda("", "vti colspan=3", "", "", "", "", "", SPROCMAIN, "si")
        Celda = Celda & sInput
        Fila = Fila & ofv.GenRow("", "", "", "", "25", "", "", Celda, "si")




        If Not rs.EOF Then

            WKFSIGOK = "S"

            ''''''''''''''''''''''''''''''''''''''''''''''''
            filaS1 = ""

            StrSql = "SELECT * FROM OBJwfldet  "
            StrSql = StrSql & " Where codobjeto=" & nCodigo
            StrSql = StrSql & " And tipobjeto = '" & stipo & "' And CKTTYPE = '" & Session("CKTTYPE") & "' order by orden,objid"
            'response.write (Chr(13))
			'response.write (StrSql & Chr(13))
			'response.write (Chr(13))
            Set RSX1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
            If Not RSX1.EOF Then

                Do While Not RSX1.EOF

                    If (RSX1(4) & "") = "NO" AND (RSX1("ESTADOPASOWKF") & "") = "S" Then WKFSIGOK = "N"

                    sInput = "RESPONSABLE"
                    MUESTRACOMBO = "NO"

                    Select Case (RSX1(3) & "")

                        Case "1"
                            sInput = "REFERENTE:"

                        Case "2"
                            sInput = "SUPERVISOR REFERENTE:"
                            MUESTRACOMBO = "SI"

                        Case "3"
                            sInput = "DUEÑO DEL PROCESO:"
                            MUESTRACOMBO = "SI"

                    End Select


                    Celda1 = ofv.gencelda("", "UT style='border-style:solid;border-width:1;border-color:#999999;' ", "", "30%", "", "", "", sInput, "si")

                    '*********************************************************************************************
                    ' SE AGREGA EL COMBO DE MOSTRAR EL PASO DEL WORKFLOW O NO, DEL SUPERVISOR O DUEÑO
                    '*********************************************************************************************
                    If MUESTRACOMBO = "SI" Then
                        scomK = "S" & Chr(9)
                        scomD = "SI" & Chr(9)
                        scomK = scomK & "N" & Chr(9)
                        scomD = scomD & "NO" & Chr(9)
                        sInput3 = ofv.gencomboFX(sForm & "AUT" & RSX1(0), RSX1("ESTADOPASOWKF").Name, RSX1("ESTADOPASOWKF"), scomK, scomD, "")
                        Celda1 = Celda1 & ofv.gencelda("", "UT style='border-style:solid;border-width:1;border-color:#999999;' ", "", "10%", "", "", "", sInput3, "si")
                    Else
                        Celda1 = Celda1 & ofv.gencelda("", "UT style='border-style:solid;border-width:1;border-color:#999999;' ", "", "10%", "", "", "", "&nbsp;", "si")
                    End If

                    If Trim(RSX1("ESTADOPASOWKF")) = "S" Then

                        sval = (RSX1(4) & "")

                        scomK = "NO" & Chr(9)
                        scomD = "SIN ASIGNAR" & Chr(9)




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



                        StrSql2 = "Select DISTINCT GRPUSR From OBJSDKDET WHERE codobjeto = 0" & SPROCMAINCOD & "  "
                        Set rsniv = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)
                        Do Until rsniv.EOF
                            '  WKUSRSEL = REPLACE(RSNIV(0) & "","USR-","")
                            '  strsql = strsql & wkfsep3 & "'" & WKUSRSEL & "'"
                            '  wkfsep3 = ","

                            xxx = getinfouser(Replace(rsniv(0) & "", "USR-", ""))
                            xnombre = Session("SGU_nombre")
                            NSELECT = NSELECT & "<OPTION VALUE='" & Replace(rsniv(0) & "", "USR-", "") & "' >" & xnombre & "</OPTION>"
                            scomK = scomK & Replace(rsniv(0) & "", "USR-", "") & Chr(9)

                            scomD = scomD & xnombre & Chr(9)



                            rsniv.MOVENEXT
                        Loop
                        Call ofv.cerrarconsulta(rsniv)



                        'NSELECT = NSELECT & "</SELECT> "


                        SINPUT2 = ofv.gencomboFX(sForm & "AUT" & RSX1(0), RSX1(4).Name, RSX1(4), scomK, scomD, "")
                        'sInput2 = "hOLA"


                        Celda1 = Celda1 & ofv.gencelda("", "UT style='border-style:solid;border-width:1;border-color:#999999;' ", "", "60%", "", "", "", SINPUT2, "si")
                        XFTROW = ofv.GenRow("", "", "", "", "25", "", "", Celda1, "si")


                    Else
                        Celda1 = Celda1 & ofv.gencelda("", "UT style='border-style:solid;border-width:1;border-color:#999999;' ", "", "60%", "", "", "", " - SIN ASIGNACIÓN - ", "si")
                        XFTROW = ofv.GenRow("", "", "", "", "25", "", "", Celda1, "si")
                    End If


                    sAccion = "asp/ActualizarEN.asp?ID=" & RSX1(0) & "&Tabla=ObjwflDET&volver=" & vuelta & "_idUpdate=0" & RSX1(0) & "_estadoactivo=" & RSX1("ESTADOPASOWKF")
                    sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & stipo
                    filaS1 = filaS1 & ofv.GenForm(sForm & "AUT" & RSX1(0), "", sAccion, "", "post", "", XFTROW, "si")

                    RSX1.MOVENEXT
                Loop

            Else

                WKFSIGOK = "N"



            End If



            CeldaT = ofv.GenTabla("", "aaa  ", "", "90% style='border-style:solid;border-width:1;border-color:#999999;' ", "", "0", "0", "1", "2", "", "", filaS1, "si")

            Celda = ofv.gencelda("", "UT COLSPAN=4 ", "CENTER", "", "", "", "", CeldaT, "si")
            Fila = Fila & ofv.GenRow("", "", "", "", "25", "", "", Celda, "si")

            ''''''''''''''''''''''''''''''''''''''''''''''''

            sForm = "wfl"

            'Texto del Mensaje
            esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:transparent;} "

            XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))

            'XMESAURL = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL

            if SESSION("URLPORT") ="N" then
            XMESAURL = TipoSite("")  & request.servervariables("SERVER_NAME")  & XMESAURL

            else
               XMESAURL = TipoSite("")  & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
            end if

            Celda = ofv.gencelda("", "ttr colspan=4", "LEFT", "", "", "", "", XMESAURL, "si")
            '  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")



            Celda = ofv.gencelda("", "ttr colspan=4", "LEFT", "", "", "", "", "MENSAJE:", "si")
            Fila = Fila & ofv.GenRow("", "", "", "", "25", "", "", Celda, "si")

            If bModificar And (CInt(sestado) = 1 Or CInt(sestado) = 6) Then
                '''   if bModificar  and (Cint(sestado) = 1 ) then

                wkfmensaje = rs(7) & ""

                If wkfmensaje = "" Then

                    If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

                        wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un CIRCUITO"
                        wkfmensaje = wkfmensaje & " pendiente de su aprobación."
                        wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
                        wkfmensaje = UCase(wkfmensaje)

                    Else

                        wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un proyecto"
                        wkfmensaje = wkfmensaje & " normativo pendiente de su aprobación."
                        wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
                        wkfmensaje = UCase(wkfmensaje)

                    End If


                    StrSql = "update OBJWFL set MSGCHAT = '" & wkfmensaje & "'  "
                    StrSql = StrSql & " Where codobjeto=" & nCodigo
                    StrSql = StrSql & " And tipobjeto = '" & stipo & "'"

                    Call ofv.crearconsultaEx(StrSql, cn, 1, volver)
                End If

                nCol = 80

                t3b = "<font class=at><textarea CLASS=DT "
                t3b = t3b & "name=" & rs(7).Name & " cols=" & nCol & " rows=6 "
                If bModificar Then
                    t3b = t3b & " onblur=" & sForm & "3.submit() "
                    t3b = t3b & " onkeypress=" & Chr(34)
                    t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
                    t3b = t3b & " extenso'); return false;}" & Chr(34) & " >"
                Else
                    t3b = t3b & " readonly>"
                End If
                t3b = t3b & wkfmensaje & "</textarea></font>"
                sInput = ofv.gencelda("", "ut colspan=4", "center", "", "", "", "", t3b, "si")
                sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
                sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & stipo
                sInput = ofv.GenForm(sForm & "3", "", sAccion, "", "post", "", sInput, "si")
            Else
                sInput = ofv.gencelda("", "ut colspan=4", "center", "", "", "", "", rs(7), "si")
            End If

            Celda = sInput
            Fila = Fila & ofv.GenRow("", "", "", "", "", "", "", Celda, "si")


        End If


        WKFSIG = "S"


        If rs(3) Then
            WKFSIG2 = "S"
        Else
            WKFSIG2 = "N"
        End If

        If rs(4) Then
            WKFMODAL = "S"
        Else
            WKFMODAL = "N"
        End If

    End If


    Call ofv.cerrarconsulta(rs)

    sHTML = sHTML & "<br><center>"
    sHTML = sHTML & ofv.GenTabla("", "aaa  ", "", "70% style='border-style:solid;border-width:1;border-color:#999999;' ", "", "0", "0", "-2", "1", "", "", Fila, "si")
    sHTML = sHTML & "</center>"

    Fila = ""

    Celda = ofv.gencelda("", "ttt colspan=4", "", "", "", "", "", "&nbsp;estado del circuito", "si")
    Formulario2 = ofv.GenRow("", "", "", "", "30", "", "", Celda, "si")

    If WKFSIG = "S" Then

        wkxhay = "N"

        Set cn5 = ofv.conectar(ofv.strconn5)

        StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & Session("TCliente") & "' and  OBJCOD = '" & Session("NCliente") & "' AND FIN = 'NO'  "
        Set rsxx = ofv.crearconsultaEx(StrSql, cn5, 1, parametros)
        If Not rsxx.EOF Then
            wkxhay = "S"

        End If
        Call ofv.cerrarconn(cn5)
        'response.write session("WKSPTPERMCT")
        if session("WKSPTPERMCE") = "" then  session("WKSPTPERMCE") = "N-S"
        wktabpermite = Split(Session("WKSPTPERMCE") & "", "-")
        wktabcircuito = UBound(wktabpermite)
        wkpermisoce = wktabpermite(wktabcircuito)




        '''  IF (Cint(sestado) = 1 or Cint(sestado) = 6 ) AND (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S"  and wkxhay <> "S" then

        'response.write " ncliente:" & session("NCliente") & "-tcliente:" & session("TCliente") & "-sestado:" & sestado & "-bmodificar:" & bmodificar & "-nomodificarce:" & session("Nomodificarce") & "-segce:" & wktabcircuito & "/" & wkpermisoce & "/" & session("WKSPTPERMCE") & "-xhay:" & wkxhay & "<br>"

        If (CInt(sestado) = 1) And (bModificar Or Session("Nomodificarce") <> "S") And wkpermisoce = "S" And wkxhay <> "S" Then

            '    Strsql = "SELECT * FROM OBJwfldet  "
            '    Strsql = Strsql & " Where codobjeto=" & nCodigo
            '    Strsql = Strsql & " And tipobjeto = '" & stipo & "' And CKTTYPE = '" & session("CKTTYPE") & "' order by orden,objid"
            '    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            '    IF RS.EOF THEN
            '       WKFSIGOK = "N"
            '    ELSE
            '       WKFSIGOK = "S"
            '    END IF


            If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

                StrSql = "SELECT * FROM ORRSKCONFIG WHERE APLIC = '" & WKAPLIC & "'  "
                Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If rs.EOF Then
                    WKFSIGOK = "N2"
                Else
                    If WKOPCION = "AUTO" Then
                        WKFCKT = rs(5)
                    Else
                        WKFCKT = rs(6)
                    End If
                End If

            Else

                StrSql = "SELECT * FROM TIPOBJETO  "
                StrSql = StrSql & " Where "
                StrSql = StrSql & " tipobjeto = '" & stipo & "' "
                Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If rs.EOF Then
                    WKFSIGOK = "N2"
                Else
                    If WKOPCION = "AUTO" Then
                        WKFCKT = rs(19)
                    Else
                        WKFCKT = rs(20)
                    End If
                End If

            End If



            If WKFSIGOK = "S" Or (WKFSIGOK = "N" And WKFSIG2 = "N" And WKOPCION = "AUTO") Then

                Celda = ""
                If WKFSIGOK = "N" Then
                    SINPUT2 = "<b>ATENCION!!</b><BR>EN CASO DE INICIAR, EL CIRCUITO CERRARA CON "
                    SINPUT2 = SINPUT2 & " <u><b>" & Session("SEGUSUARIO") & "</b></u>  COMO UNICO AUTORIZANTE.<BR>"
                Else
                    SINPUT2 = "<b>ATENCION!!</b><BR>EL CIRCUITO ESTA EN CONDICIONES DE SER INICIADO POR "
                    SINPUT2 = SINPUT2 & " <u><b>" & Session("SEGUSUARIO") & "</b></u>  COMO ORIGINANTE.<BR>"
                End If
                Celda = ofv.gencelda("", "uti colspan=3 style='text-align:justify;' ", "", "", "", "", "", SINPUT2, "si")

             '' Agregado 20042015
             '''sInput = ofv.GenerarInput("boton1", "INICIAR", "submit", "edPropiedades2", "5", "bt", "")
                sInput = ofv.GenerarInput("boton1", "INICIAR", "submit", "edPropiedades2", "5", "bt onclick=""this.style.cursor='wait';this.style.visibility='hidden';""","")
	         '' Agregado 20042015 FIN


                Celda = Celda & ofv.gencelda("", "uti", "center", "10%", "", "", "", sInput, "si")

                Fila = ofv.GenRow("", "", "", "", "", "", "", Celda, "si")

                sAccion = "EDWKFSTART.asp?OPCION=" & WKOPCION & "&MODAL=" & WKFMODAL & "&" & request.querystring
                Formulario2 = Formulario2 & ofv.GenForm("edPropiedades2", "", sAccion, "", "post", "", Fila, "si")

            Else

                Celda = ""
                If WKFSIGOK = "N" Then
                    sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, NO INFORMO LA LISTA DE AUTORIZANTES.<BR>"
                Else
                    sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, EL TIPO DE DOCUMENTO NO ESTA CONFIGURADO.<BR>"
                End If
                Celda = ofv.gencelda("", "uti colspan=3 style='text-align:justify;' ", "", "", "", "", "", sInput, "si")

                sInput = "&nbsp;"
                Celda = Celda & ofv.gencelda("", "ut", "center", "5%", "", "", "", sInput, "si")

                Fila = ofv.GenRow("", "", "", "", "", "", "", Celda, "si")

                Formulario2 = Formulario2 & Fila

            End If

        Else

            Celda = ""

            If (CInt(sestado) = 1) And (bModificar Or Session("Nomodificarce") <> "S") And wkpermisoce = "S" And wkxhay = "S" Then
                sInput = "<FONT COLOR=#005500>EN CURSO.</FONT><BR>"
            Else
                sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO DESDE AQUI O NO ESTA AUTORIZADO PARA INICIARLO.<BR>"
            End If
            Celda = ofv.gencelda("", "uti colspan=3 style='text-align:justify;' ", "", "", "", "", "", sInput, "si")

            sInput = "&nbsp;"
            Celda = Celda & ofv.gencelda("", "ut", "center", "5%", "", "", "", sInput, "si")

            Fila = ofv.GenRow("", "", "", "", "", "", "", Celda, "si")

            Formulario2 = Formulario2 & Fila


        End If

        ''     IF (Cint(sestado) = 1 ) AND (bModificar   or session("Nomodificarce") <> "S") and wkpermisoce = "S"   and wkxhay <> "S"  then

        ''   response.write " ncliente:" & session("NCliente") & "-tcliente:" & session("TCliente") & "-sestado:" & sestado & "-bmodificar:" & bmodificar & "-nomodificarce:" & session("Nomodificarce") & "-segce:" & wktabcircuito & "/" & wkpermisoce & "/" & session("WKSPTPERMCE") & "-xhay:" & wkxhay & "<br>"

        If ((CInt(sestado) = 2 Or CInt(sestado) = 5 Or CInt(sestado) = 6) Or (CInt(sestado) = 1 And wkxhay = "S")) And wkpermisoce = "S" Then

            '  sInput = ofv.GenerarInput("boton1","Consultar historial","submit","edPropiedades3","5","bt","")
            '  sAccion = "CoHistWkf.asp"
            '  sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
            '  Celda = ofv.gencelda("","Vt colspan=2","center","30%","","","",sInput,"si")

            If (bModificar) And wkpermisoce = "S" Then
                sInput = ofv.GenerarInput("boton1", "Cancelar Circuito", "submit", "edPropiedades3", "5", "bt style='background-color:#993333;' ", "")
                sAccion = "EDCANCELWkf.asp?OPCION=CONF"
                sInput = ofv.GenForm("edPropiedades3", "", sAccion, "", "post", "", sInput, "si")
            Else
                sInput = "&nbsp;"

            End If


            Celda = ""

            Celda = Celda & ofv.gencelda("", "Vt colspan=2 ", "center", "30%", "", "", "", sInput, "si")

            sInput = "&nbsp;"
            Celda = Celda & ofv.gencelda("", "Vt colspan=2", "center", "30%", "", "", "", sInput, "si")

            Fila = ofv.GenRow("", "", "", "", "", "", "", Celda, "si")


            Formulario2 = Formulario2 & Fila


        End If




    End If

    'xxxxxxxxxx



    If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

        sAccion = Session("ASPSTR") & "&ACCIONEJEC=ABREESTAD"
        Retornar = ofv.BotonVolver(sAccion)


    End If


    sHTML = sHTML & "<br><center>"
    sHTML = sHTML & ofv.GenTabla("", "aaa style='border-style:solid;border-width:1;border-color:#999999;' ", "", "70% ", "", "0", "0", "1", "0", "", "", Formulario2, "si")
    sHTML = sHTML & "</center>"



    Call ofv.cerrarconsulta(rs)
    Call ofv.cerrarconn(cn)
    sestado = 1

	'--------------------------------------------------------------------------------
	  sHTML = sHTML & "</div>"

	  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	  Fila = ofv.GenRow("","","","","","","",Celda,"si")
	  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	  'Estructura del Indice

	  sHTML = sHTML & CorPnlIndx("N","")

sHTML = sHTML & ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "no", Agregar, Retornar, "")

'--------------------------------------------------------------------------------
    sHTML = sHTML & "</body></html>"
    '' end if
End If

If (CInt(sestado) = 1 Or CInt(sestado) = 6) Then response.write sHTML

Function GENPROCSECUENEX(scodmaster)

    GENPROCSECUENEX = ""

    StrSql = "Select * From CSRELPATHDESCRI  where COBJ = " & scodmaster
    '' response.write strsql & "<br>"
    Set WRS = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not WRS.EOF Then

        WKITOPE = 8

        For WKI = 0 To 7

            WKIPROC = WRS(WKI) & ""

            If WKIPROC <> "" Then
                WKITOPE = WKI
                Exit For
            End If


        Next

        WKICCOL = 7 - WKITOPE

        WKINCOL = 0

        FilaW = ""

        XCelda = ""
        XCelda2 = ""

        SPROCMAIN = WRS(WKITOPE)
        SPROCMAINCOD = WRS(WKITOPE + 8)

        For WKI = WKITOPE To 7




            XCelda = XCelda & XCelda2 & WRS(WKI)

            XCelda2 = "<b><font color='#00aa00' size=2 >&#0124;</font></b>"



        Next


        If WKITOPE < 8 Then

            GENPROCSECUENEX = XCelda & XCelda2

        End If



    End If


End Function

%>
