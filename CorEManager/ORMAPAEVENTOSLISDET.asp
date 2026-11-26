<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = ofv.CheckUsuario()
If ok <> "Y" Then
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    sHTML = ok
Else
    Response.expires = 0


    CantRegMostrarx = 6
    Cantregmoverx = CDbl(request.querystring("CantRegMover"))
    vuelta = ofv.vueltaasp("../", CantRegMostrarx, Cantregmoverx)
    smodal = request.querystring("MODAL")
    sMODALMAIN = request.querystring("MODALMAIN")
    SCODTRACER = request.querystring("CODTRACER")
    sFECHA = request.querystring("FECHA")
    sOPCION = request.querystring("OPCION")



    If smodal = "" Then smodal = "CSN"

    sHTML = ofv.FormHeader(SESSION("FormN"))

    WKVISION = ""


    WSPSESION = "ORWHERE"

    ofv.ObtenerAtributos SESSION("Form"), ""
    bAgregar = ofv.AAgregar
    bModificar = ofv.AModificar
    bSuprimir = ofv.AEliminar
    bConsultar = ofv.AConsultar

    Set cn = ofv.conectar(ofv.strconn2)


    SORLINNEG = request.querystring("ORLINNEG")
    SORTIPORG = request.querystring("ORTIPORG")

    SORLINNEGW = ""
    SORTIPORGW = ""

    SCONECTOR = " AND "

    WSPWHERE = SESSION(WSPSESION)


    If SORLINNEG <> "TODOS" Then

        If SORLINNEG = "0" Then

            SORLINNEGW = " ((C.NIVEL1 = '0') OR (C.NIVEL1 IS NULL) ) "

        Else

            SORLINNEGW = " ((C.NIVEL1 = '" & SORLINNEG & "') OR (D.NIVEL1 = '" & SORLINNEG & "'  ) )"

        End If

        WSPWHERE = WSPWHERE & SCONECTOR & SORLINNEGW


    End If

    If SORTIPORG <> "TODOS" Then

        If SORTIPORG = "0" Then

            SORTIPORGW = " ((A.TRCEVENTORIESGO1 = '0') OR (A.TRCEVENTORIESGO1 IS NULL) ) "

        Else

            SORTIPORGW = " (A.TRCEVENTORIESGO1 = '" & SORTIPORG & "') "

        End If

        WSPWHERE = WSPWHERE & SCONECTOR & SORTIPORGW

    End If







    'RESPONSE.WRITE smodal & "<BR><BR>"






    If smodal = "REG" Or smodal = "ENR" Or smodal = "ENCA" Or smodal = "CSN" Then




        StrSql = "     select A.*, B.ACTIVIDAD, "

        StrSql = StrSql & "  CASE A.TRCESTADO "
        StrSql = StrSql & "        WHEN 'CAN' THEN '4' "
        StrSql = StrSql & "        WHEN 'REG' THEN '1' "
        StrSql = StrSql & "        WHEN 'ENR' THEN '2' "
        StrSql = StrSql & "        WHEN 'ENRF' THEN '2' "
        StrSql = StrSql & "        WHEN 'FIN' THEN '3' "

        StrSql = StrSql & "         END AS ESTORD, C.NIVEL1 , A.TRCEVENTORIESGO1, A.TRCPORCDISTRIB, D.NIVEL1 AS NIVEL1_2, D.PORCENTAJE, C.DESCRIPCION as DESCRIPCION_1, D.DESCRIPCION as DESCRIPCION_2"
		'StrSql = StrSql & "         END AS ESTORD, C.NIVEL1 , A.TRCEVENTORIESGO1, A.TRCDISTRIB, A.TRCPORCDISTRIB, D.PORCENTAJE, C.DESCRIPCION as DESCRIPCION_1, D.DESCRIPCION as DESCRIPCION_2"
		
		Strsql = Strsql & " ,CASE A.TRCESTADO WHEN 'NO' THEN 'RESGUARDO RECTIFICATIVA' "
	 Strsql = Strsql & " ELSE ( SELECT DESCRIPCION FROM ORESTADOS WHERE CODIGO = A.TRCESTADO ) "
	 Strsql = Strsql & " END AS ESTADODSC "


        StrSql = StrSql & " "
        StrSql = StrSql & " FROM ORTRACER A "
        StrSql = StrSql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
        StrSql = StrSql & " LEFT JOIN ORLINNEGO2  C ON (A.TRCNEGOLINE = C.OBJID)  "
        StrSql = StrSql & " LEFT JOIN (SELECT D2.NIVEL1, D2.DESCRIPCION, D1.* FROM ORDISTRIBLN D1 INNER JOIN ORLINNEGO2 D2 ON (D1.LINNEGO = D2.OBJID)) "
        StrSql = StrSql & "      D ON (A.CODTRACER = D.CODTRACER AND A.TRCDISTRIB = 'S')  "
        StrSql = StrSql & "   " & WSPWHERE & "  ORDER BY    A.OBJID, ESTORD  "
        WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

        If smodal = "ENCA" Then WKACCION = " - ENCADENAMIENTO DE EVENTOS DE PERDIDA"
        If smodal = "CSN" Then WKACCION = " - CONSULTA DE EVENTOS DE PERDIDA"

		'response.write chr(13) & chr(13)
        'response.write strsql
		'response.write chr(13) & chr(13)

        Celda = ""

        wkproccod = "&nbsp;"

        Celda = Celda & ofv.gencelda("", "uti  ", "", "90%", "", "", "", "BASE DE EVENTOS DE RIESGO " & WKACCION, "si")
        Celda = Celda & ofv.gencelda("", "ut   ", "", "2%", "", "", "", wkproccod, "si")
        Fila = ofv.genrow("", "", "", "", "", "", "", Celda, "si")

        sHTML = sHTML & ofv.GenTabla("", "TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ", "", "100% ", "30", "0", "", "-1", "1", "", "", Fila, "si")

        Celda = ""
        Celda = Celda & ofv.gencelda("", "uti  ", "", "10%", "", "", "", "SELECCION: ", "si")
        Celda = Celda & ofv.gencelda("", "ut   ", "", "90%", "", "", "", WSPWHERE, "si")
        Fila = ofv.genrow("", "", "", "", "", "", "", Celda, "si")

        ' sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;BORDER-COLOR:#993333;' valign='middle'  ","","100% ","30","0","","-1","1","","",Fila,"si")
        sHTML = sHTML & "<br>"

        'Encabezados

        Formulario = ""
        FilaX = ""
        Celda = ""


        ' response.write strsql & "<br>"
        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If rs.EOF Then


        Else
            posi = 0

            xTrColor = "#f0f0f0"
            '''response.write strsql & "<br>"

            '' if not rs.eof then rs.move(CantRegMoverx)
            '' posi = 0
            '' Do While (Not rs.EOF) and (posi < CantRegMostrarx)

            WKIMPMONTOFFF = 0

            'response.write rs(16).name & "---" & rs(45).name
            Do While (Not rs.EOF)


                sForm = "form" & posi


                FilaX = ""

                Celda = ""

                'CODIGO
                Celda = Celda & ofv.gencelda("", "VTI STYLE='background-color:#888888;color:#ffffff;' ", "", "5%", "", "", "", "CODIGO", "si")


                'INCIDENTE
                Celda = Celda & ofv.gencelda("", "VTI STYLE='background-color:#888888;color:#ffffff;' ", "", "30%", "", "", "", "FECHA", "si")

                'CARACTER
                Celda = Celda & ofv.gencelda("", "VTI STYLE='background-color:#888888;color:#ffffff;'", "", "30%", "", "", "", "CARACTER", "si")

                'CUENTA
                Celda = Celda & ofv.gencelda("", "VTI STYLE='background-color:#888888;color:#ffffff;'", "", "30%", "", "", "", "CUENTA", "si")

                ''  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



                Celda = ""


                'CODIGO
                sInput = rs(0)

                sAccion = "ORBASERIESGOSADD.ASP?MODAL=CSN&CODTRACER=" & rs(1)

                Celda = Celda & ofv.gencelda("", "vT onclick='document." & sForm & ".submit();' title='ver evento.' style='cursor:pointer;background-color:#F7AA22;border-style:outset;border-width:2;' ROWSPAN=3  ", "center", "", "", "", "", sInput, "si")


                '''   Celda = Celda & ofv.gencelda("","vT  style='cursor:pointer;background-color:#F7AA22;border-style:outset;border-width:2;' ROWSPAN=3  ","center","","","","",sInput,"si")


                'F.INCIDENTE
                If CLng(rs(3)) = 0 Then
                    dFechan = "-------"
                Else
                    dFechan = GTFechadorSH(rs(3))
                End If
                sInput = dFechan
                Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "5%", "", "", "", "FECHA:&nbsp;", "si")
                Celda = Celda & ofv.gencelda("", "UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "25%", "", "", "", sInput, "si")



                'CARACTER

                sInputx = "-------"

                StrSql = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs(24)
                Set rsniv = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If Not rsniv.EOF Then sInputx = rsniv(0)
                Call ofv.cerrarconsulta(rsniv)

                sInput = sInputx
                Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "CARACTER:&nbsp;", "si")
                Celda = Celda & ofv.gencelda("", "UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput, "si")



                'CUENTA CONTABLE
                If (rs(29) & "") = "" Then
                    sInput = "-------"
                Else
                    sInput = rs(29)
                End If
                Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "CUENTA:&nbsp;", "si")
                Celda = Celda & ofv.gencelda("", "UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput, "si")


                If smodal = "ENCA" Then
                    sInput = "<A CLASS=CC HREF='ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=ENR&OPCION=SI&PRETRACER=" & rs(1) & "' >RELACIONAR</A>"
                    Celda = Celda & ofv.gencelda("", "vT  title='Relacionar este evento.' style='cursor:pointer;background-color:#ddddff;border-style:outset;border-width:2;' ROWSPAN=3  ", "center", "", "", "", "", sInput, "si")
                End If

                FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")




                Celda = ""
                'CLIENTE
                Celda = Celda & ofv.gencelda("", "VTI STYLE='background-color:#888888;color:#ffffff;' ", "", "", "", "", "", "CLIENTE", "si")

                'MONTO
                Celda = Celda & ofv.gencelda("", "VTI STYLE='background-color:#888888;color:#ffffff;'", "", "", "", "", "", "MONTO", "si")

                'ESTADO
                Celda = Celda & ofv.gencelda("", "VTI STYLE='background-color:#888888;color:#ffffff;'", "", "", "", "", "", "ESTADO", "si")

                '  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



                Celda = ""
                'CLIENTE
                If (rs(28) & "") = "" Then
                    sInput = "-------"
                Else
                    sInput = rs(28)
                End If
                Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "CLIENTE:&nbsp;", "si")
                Celda = Celda & ofv.gencelda("", "UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput, "si")


                'MONTO
				IF ( isNumeric(rs(16)) ) THEN
					If CLng(rs(16)) = 0 Then
						sInput = "-------"
					Else
						sInput = rs(16)
					End If
				ELSE
					sInput = "-------"
				END IF
                Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "MONTO NETO:&nbsp;", "si")
                Celda = Celda & ofv.gencelda("", "UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput, "si")


                'ESTADO
                'wkestado = rs(17) & ""
                'Select Case wkestado
                '   Case "CAN"
                '        wkestadoD = "<font color='#888888' >CANCELADO</font>"
                '    Case "REG"
                '        wkestadoD = "<font color='#993333' >REGISTRADO</font>"
                '    Case "ENR"
                '        wkestadoD = "<font color='#009900' >ENRIQUECIDO</font>"
                '    Case "ENRF"
                '        wkestadoD = "<font color='#009900' >PENDIENTE</font>"
                '   Case "FIN"
                '       wkestadoD = "<font color='#000000' ><B>FINALIZADO</B></font>"
                '    Case Else
                '        wkestadoD = "<font color='#000099' >EN RESGUARDO</font>"
                'End Select
                'sInput = wkestadoD
				
				'ESTADO
				wkestado = rs("TRCESTADO") & ""
				wkmodoOT = rs("TRCMARCAACTP") & ""

				wkestadoD = rs("ESTADODSC") & ""

				IF wkestado = "NO" AND wkmodoOT = "RM" THEN
					wkestadoD = "<font color='#000000' >RESGUARDO RECTIFICATIVA</font>"
				ELSEIF (wkestadoD & "") = "" THEN
					wkestadoD = "<font color='#000099' >EN REVISION</font>"
				ELSE
					wkestadoD = "<font color='#009900' >" & wkestadoD & "</font>"
				END IF

                Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "ESTADO:&nbsp;", "si")
                Celda = Celda & ofv.gencelda("", "uTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", wkestadoD, "si")

                FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

                Celda = ""
                'OBSERVACIONES
                sInput = rs(11) & "&nbsp;"
                Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "DESCRIPCION:&nbsp;", "si")
                Celda = Celda & ofv.gencelda("", "UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput, "si")

                FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

                posi = posi + 1

                WKVINODISTRIB = "S"
                WKCODDISTRIB = rs(0)

                ''''   IF (RS(44) & "") <> "S" THEN

                If (rs("TRCDISTRIB") & "") <> "S" Then

                    WKIMPMONTOFFF = WKIMPMONTOFFF + Round(CDbl(rs(16)) * 1, 3)

                    Celda = ""
                    'RESUMEN

                    Celda = Celda & ofv.gencelda("", "UT style='background-color:#993333;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "&nbsp;", "si")
                    sInput = rs("DESCRIPCION_1") & "&nbsp;"
                    Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "LINEA:&nbsp;", "si")
                    Celda = Celda & ofv.gencelda("", "UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput, "si")

                    sInput = "100%&nbsp;"
                    Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "PORCENTAJE:&nbsp;", "si")
                    Celda = Celda & ofv.gencelda("", "UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "RIGHT", "", "", "", "", sInput, "si")

                    sInput = "$&nbsp;" & FormatNumber(rs(16), 2, 0, 0, -1) & "&nbsp;"
                    Celda = Celda & ofv.gencelda("", "UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#888888;' ", "", "", "", "", "", "TOTAL:&nbsp;", "si")
                    Celda = Celda & ofv.gencelda("", "UTi style='color:#000000;border-style:SOLID;border-width:1;border-color:#888888;' ", "RIGHT", "", "", "", "", sInput, "si")

                    FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

                    WKVINODISTRIB = "N"
                    rs.Movenext

                Else

                    Celda = ""
                    'RESUMEN

                    Celda = Celda & ofv.gencelda("", "UT style='background-color:#888888;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "&nbsp;", "si")

                    Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "LINEA:&nbsp;", "si")
                    Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "PORCENTAJE:&nbsp;", "si")
                    Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "MONTO:&nbsp;", "si")
                    FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

                    WKIMPRI = "S"
                    WKIMPLINNEG = ""
                    WKIMPPORCEN = ""
                    WKIMPMONTOP = ""
                    WKIMPMONTOf = 0

                    If (rs("NIVEL1") & "") = (SORLINNEG & "") Then

                        WKIMPRI = "S"
                        WKIMPLINNEG = rs("DESCRIPCION_1")
                        WKIMPPORCEN = CDbl("0" & rs("TRCPORCDISTRIB"))
                        WKIMPMONTOP = Round(CDbl(rs(16)) * CDbl(WKIMPPORCEN) / 100, 3)
                        WKIMPMONTOf = WKIMPMONTOf + WKIMPMONTOP
                        WKIMPMONTOFFF = WKIMPMONTOFFF + WKIMPMONTOP
                    Else
                        WKIMPRI = "N"
                    End If


                    If WKIMPRI = "S" Then

                        Celda = ""
                        'RESUMEN

                        Celda = Celda & ofv.gencelda("", "UT style='background-color:#6699CC;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "&nbsp;", "si")
                        sInput = WKIMPLINNEG & "&nbsp;"
                        Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput , "si")

                        sInput = WKIMPPORCEN & "%&nbsp;"
                        Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "RIGHT", "", "", "", "", sInput , "si")

                        sInput = "$&nbsp;" & FormatNumber(WKIMPMONTOP, 2, 0, 0, -1) & "&nbsp;"
                        Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "RIGHT", "", "", "", "", sInput , "si")

                        FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

                    End If




                    Do Until (WKVINODISTRIB = "N")

                        If (rs("NIVEL1_2") & "") = (SORLINNEG & "") Then
                            WKIMPRI = "S"
                            WKIMPLINNEG = rs("DESCRIPCION_2")
                            'IF (ISNUMERIC(rs("PORCENTAJE"))) THEN
								WKIMPPORCEN = rs("PORCENTAJE")
							'ELSE
							'	WKIMPPORCEN = 1
							'END IF
							
                            If (WKIMPPORCEN & "") = "0" Then
                                StrSql = " select sum(porcentaje) from ordistribln where codtracer = '" & rs(1) & "' group by codtracer"
                                ''    response.write strsql & "<br>"
                                Set rsx1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                                If Not rsx1.EOF Then
                                    WKIMPPORCEN = 100 - (CDbl("0" & rsx1(0)) + CDbl("0" & rs("TRCPORCDISTRIB")))
                                    '  WKIMPPORCEN = 100 - (cdbl( rsx1(0)) + cdbl( RS(54)))
                                End If

                            End If


                            WKIMPMONTOP = Round(CDbl(rs(16)) * CDbl(WKIMPPORCEN) / 100, 3)
                            WKIMPMONTOf = WKIMPMONTOf + WKIMPMONTOP
                            WKIMPMONTOFFF = WKIMPMONTOFFF + WKIMPMONTOP
                        Else
                            WKIMPRI = "N"
                        End If


                        If WKIMPRI = "S" Then

                            Celda = ""
                            'RESUMEN

                            Celda = Celda & ofv.gencelda("", "UT style='background-color:#6699CC;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "&nbsp;", "si")
                            sInput = WKIMPLINNEG & "&nbsp;"
                            Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", sInput, "si")

                            sInput = WKIMPPORCEN & "%&nbsp;"
                            Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "RIGHT", "", "", "", "", sInput, "si")

                            sInput = "$&nbsp;" & FormatNumber(WKIMPMONTOP, 2, 0, 0, -1) & "&nbsp;"
                            Celda = Celda & ofv.gencelda("", "UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "RIGHT", "", "", "", "", sInput, "si")

                            FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

                        End If



                        rs.Movenext
                        If Not rs.EOF Then
                            WKCODDISTRIBnew = (rs(0) & "")
                        Else
                            WKCODDISTRIBnew = "FINDATOS"
                        End If
                        '' response.write " rs0=" & WKCODDISTRIBnew  & " WKCODDISTRIB:" & (WKCODDISTRIB & "")  & " eof:" & rs.eof & "<br>"


                        If WKCODDISTRIBnew <> (WKCODDISTRIB & "") Then
                            WKVINODISTRIB = "N"

                            Celda = ""
                            'RESUMEN

                            Celda = Celda & ofv.gencelda("", "UT style='background-color:#993333;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "&nbsp;", "si")

                            sInput = "TOTAL&nbsp;"
                            Celda = Celda & ofv.gencelda("", "UT COLSPAN=4 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#888888;' ", "RIGHT", "", "", "", "", sInput, "si")

                            sInput = "$&nbsp;" & FormatNumber(WKIMPMONTOf, 2, 0, 0, -1) & "&nbsp;"
                            Celda = Celda & ofv.gencelda("", "UT COLSPAN=2 style='color:#000000;border-style:SOLID;border-width:1;border-color:#888888;' ", "RIGHT", "", "", "", "", sInput, "si")

                            FilaX = FilaX & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")





                        End If

                    Loop
                End If



                Formulario = Formulario & FilaX





                sInput = "&nbsp;"
                Celda = ofv.gencelda("", "UT colspan=7 style='border-TOP-style:SOLID;border-TOP-width:2;border-TOP-color:#000000;' ", "", "", "", "", "", sInput, "si")

                FilaX = ofv.genrow("", "AAA bgcolor='#fff9f9' ", "", "", "25", "", "", Celda, "si")


                '''''   Formulario = Formulario & Filax



                Formulario = Formulario & ofv.GenForm(sForm, "", sAccion, "_NEW", "post", "", FilaX, "si")






            Loop

            Celda = ""
            'RESUMEN

            Celda = Celda & ofv.gencelda("", "UT style='background-color:#000000;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "", "", "", "", "", "&nbsp;", "si")

            sInput = "TOTAL&nbsp;GENERAL&nbsp;"
            Celda = Celda & ofv.gencelda("", "UT COLSPAN=4 style='color:#000080;background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "RIGHT", "", "", "", "", sInput, "si")

            sInput = "$&nbsp;" & FormatNumber(WKIMPMONTOFFF, 2, 0, 0, -1) & "&nbsp;"
            Celda = Celda & ofv.gencelda("", "VT COLSPAN=2 style='color:#000080;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ", "RIGHT", "", "", "", "", sInput, "si")

            Formulario = Formulario & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")


        End If
        sHTML = sHTML & ofv.GenTabla("", "", "", "98% ", "", "0", "", "0", "0", "", "", Formulario, "si")

        'Botonera
        Agregar = ""
        Buscar = ""
        Celda = ""

        WSPREEMP = "modal=" & smodal & "&CODTRACER=" & SCODTRACER & "&ORLINNEG=" & SORLINNEG & "&ORTIPORG=" & SORTIPORG & "&"

        '''  sHTML = sHTML &  replace(ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda),"cantregmostrar=",WSPREEMP & "cantregmostrar=")

        sHTML = sHTML & ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "", "", "", "")

        Call ofv.cerrarconsulta(rs)


    Else



    End If


    Call ofv.cerrarconn(cn)

    sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML



%>
