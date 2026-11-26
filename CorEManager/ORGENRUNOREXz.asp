
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
RESPONSE.buffer = True
ok = ofv.CheckUsuario()
If ok <> "Y" Then
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    sHTML = ok
Else
    RESPONSE.expires = 0
    server.scripttimeout = 60 * 20
    Set cn = ofv.conectar(ofv.strconn2)

    CantRegMostrarx = 8
    CantRegMoverX = CDbl(REQUEST.QUERYSTRING("CantRegMover"))
    vuelta = ofv.vueltaasp("../", CantRegMostrarx, CantRegMoverX)
    Dim ids(2, 1000)
    Dim IDsEx(2, 1)
    Dim idspOS
    Dim POSI
    Dim CContable
    Dim primerReg
    Dim totalmontoRepet

    Dim TeventosRepetitivos
    Dim TeventosEncadenados
    Dim mes
    Dim fecha1
    Dim Ahijos(500)
    Dim aRecordSave(100, 10)
    Dim aRecordSave2(100, 0)

    Dim TOPEMONTO
    Dim ENCAMENORES
    Dim REPESUMA
    Dim SORFECINIHW1

    Dim REGVA
    primerReg = True
    sopcion = REQUEST.QUERYSTRING("opcion") & ""
    idspOS = 0
    Dim EVENTORAIZ
    EVENTORAIZ = 1
    ULTEVENTORAIZ = 0
    PRIEVENTORAIZ = 0
    TeventosRepetitivos = 0
    TeventosEncadenados = 0
    totalmontoRepet = 0
    SORFECINIHW1 = ""
    TOPEMONTO = 0

    strimestre = 0
    stANO = 0

    'strimestre= request.FORM("TRIMESTRE")

    ORFECHA = REQUEST.Form("ORFECHA")

    If Not IsDate(ORFECHA) Then
        SX = GetFecTrim(strimestre, Year(Date), feciniT, fecfinT)
        feciniT = CLng(CDate(feciniT)) - CLng(CDate("01-01-1900"))
        fecfinT = CLng(CDate(fecfinT)) - CLng(CDate("01-01-1900"))
    Else
        iTRIM = GetTrimestre(CDate(ORFECHA))
        SX = GetFecTrim(iTRIM, Year(CDate(ORFECHA)), feciniT, fecfinT)
        feciniT = CLng(CDate(ORFECHA)) - CLng(CDate("01-01-1900"))
        fecfinT = CLng(CDate(fecfinT)) - CLng(CDate("01-01-1900"))
    End If


    LOTE = REQUEST.QUERYSTRING("LOTE")

    STRSQL = "SELECT * FROM ORCONFIGBCRA"
    Set rs = ofv.crearconsultaEx(STRSQL, cn, 1, volver)
    If Not rs.EOF Then
        FCARGA = rs("FECREGIS")
        FDESCUBRIMIENTO = rs("FECIDENT")
        FORIGEN = rs("FECORIG")
        FCONTABLE = rs("FECCONTA")
        FFINAL = rs("FECFINAL")
        TOPEMONTO = CDbl(rs("TOPEMONTO"))
        ENCAMENORES = rs("ENCAMENORES")
        REPESUMA = rs("REPESUMA")
    End If

    '  response.write TOPEMONTO
    Call ofv.CERRARCONSULTA(rs)


    STRSQL = "DELETE FROM ORRIERRLOTE WHERE CODLOTE='" & LOTE & "'"
    Set rs = ofv.crearconsultaEx(STRSQL, cn, 1, volver)


    STRSQL = "SELECT * FROM ORREGINFOBCRA WHERE CODLOTE='" & LOTE & "'"
    Set rs = ofv.crearconsultaEx(STRSQL, cn, 1, volver)
    If Not rs.EOF Then
        strimestre = rs("TRIMESTRE")
        stANO = rs("ANO")
        SX = GetFecTrim(strimestre, rs("ANO"), feciniT, fecfinT)
        feciniT = CLng(CDate(feciniT)) - CLng(CDate("01-01-1900"))
        fecfinT = CLng(CDate(fecfinT)) - CLng(CDate("01-01-1900"))
        ScodDiseno = rs("coddiseno")
        Scodentidad = rs("codentidad")
    End If
    Call ofv.CERRARCONSULTA(rs)



    STRSQL = "DELETE FROM ORREGINFOBCRADET WHERE codLOTE='" & LOTE & "'"
    '   response.write STRSQL & "<br>"
    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, volver)


    STRSQL = ""
    STRSQL = STRSQL & " SELECT ULTEVENTO FROM ORREGINFOBCRA "
    STRSQL = STRSQL & " where (ano = '0" & stANO & "' and trimestre < '0" & strimestre & "' ) or (ano < '0" & stANO & "' ) "
    STRSQL = STRSQL & " order by ano desc, trimestre desc, estado desc, objid desc "
    Set rs = ofv.crearconsultaEx(STRSQL, cn, 1, volver)
    If Not rs.EOF Then
        EVENTORAIZ = CLng("0" & rs(0))
        PRIEVENTORAIZ = EVENTORAIZ
        EVENTORAIZ = EVENTORAIZ + 1
    End If










    sHTML = ofv.FormHeader(SESSION("FormN"))


    '''  RESPONSE.WRITE STRSQL & "<BR><BR><BR><BR>"

    Formulario = ""





    'Encabezados
    '  Formulario = ofv.Encabezados("ID EVENTO","COD","ENTIDAD","F.INF","EV RAIZ","SECU","CARACTER","F.INI","F.DESC","F.REG.CONT","F.FIN","DESC","T.EVEN","L.NEGO","SUC","PROCESO","PROD/SERV","MONTO","PART.CONTABLE","COB.POL","VINC.RIESGO","AERA/SOP")
    'response.write strsql











    TOTALREGBRUTOMENOR = CDbl("0")
    TOTALBRUTOMENOR = CDbl("0")

    TOTALREGBRUTOMAYOR = CDbl("0")
    TOTALBRUTOMAYOR = CDbl("0")

    TOTALENVIA = 0
    TOTALNOENVIA = 0

    WSPSESION = "SWHERERUNOR"
	SESSION(WSPSESION) = ""
	WKSETTABLA = "ORTRACER"
	SCONECTOR = ""




    ' aca procesamos rectificatorias TRCMARCASEND=R




    ' aca procesamos normales TRCMARCASEND=N  ajustes    TRCMARCASEND=J




    STRSQL = SqlEventosNormales()
 '   response.write strsql
    Set rs = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
    x = EventosNormales(rs)


    sHTML = sHTML & ofv.GenTabla("", "", "", "80% ", "", "0", "", "0", "0", "", "", Formulario & x, "si")

    sAccion = "ORGENRUNOR.asp?CantRegMover=" & CantRegMoverX & ""
    Celda = ofv.BotonVolver(sAccion)

    STRSQL = "SELECT COUNT(*),sum(ABS(TRCIMPBRUTO)) FROM ORTRACER WHERE NOT (ortracer.codTRACER like 'DOBLIGA%') AND ("





    'agrego  FCARGA FDESCUBRIMIENTO etc
    SORFECINIHW = SCONECTOR & "  (ORTRACER.TRCFECACTP > = " & feciniT & " and ORTRACER.TRCFECACTP < = " & fecfinT & ") )"
    STRSQL = STRSQL & SESSION(WSPSESION) & SORFECINIHW    ' & SORFECINIHW1


    ' RESPONSE.WRITE STRSQL & "<BR>"
    Set RST = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)



    TOTAL = RST(0)
    TOTALMONTOTRI = Replace(RST(1) & "", ",", ".")
    Call ofv.CERRARCONSULTA(RST)



    ' todos los que el raiz no esta en el trimestre y los relacionados si ,entonces resto al total

    STRSQL = "SELECT b.objid ,ABS(b.TRCIMPBRUTO) FROM ORTRACER  "

    STRSQL = STRSQL & " inner join ORTRACER  b  on (ORTRACER.codtracer=b.trccodtracerp and (" & SORFECINIHW1 & " ))"

    STRSQL = STRSQL & " where  NOT (ortracer.codTRACER like 'DOBLIGA%') and ORTRACER.trccodtracerp is null and not(" & SESSION(WSPSESION) & " )"
    '   RESPONSE.WRITE strsql




    Set RST = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
    '




    Do While Not RST.EOF

        '   TOTAL = total + 1
        '   TOTALMONTOTRI = TOTALMONTOTRI + REPLACE(RST(1)&"", ",", ".")

        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
        STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA','"&rst(0) &"','1004')"
        Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
        '  RESPONSE.WRITE STRSQL & "<BR>"
        RST.MOVENEXT
    Loop




    Call ofv.CERRARCONSULTA(RST)





    StrSql2 = "SELECT COUNT(*) FROM ORREGINFOBCRADET WHERE CODLOTE='" & LOTE & "'"
    Set RST2 = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)

    If Not RST2.EOF Then
        TX = RST2(0)
    Else
        TX = 0
    End If
    Call ofv.CERRARCONSULTA(RST2)

    StrSql3 = "SELECT COUNT(A.Z) FROM "
    StrSql3 = StrSql3 & " (SELECT SUM(1) AS Z FROM ORRIERRLOTE "
    StrSql3 = StrSql3 & "WHERE CODLOTE='" & LOTE & "'"
    StrSql3 = StrSql3 & "GROUP BY  NROEVENTO) A"
    '   response.write STRSQL3
    Set RST2 = ofv.crearconsultaEx(StrSql3, cn, 1, parametros)
    If Not RST2.EOF Then
        TX2 = RST2(0)
    Else
        TX2 = 0
    End If



    StrSql2 = "SELECT COUNT(*) FROM ORREGINFOBCRADET WHERE CODLOTE='" & LOTE & "'"
    Set RST2 = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)

    If Not RST2.EOF Then
        TTOTAL = RST2(0)
    Else
        TTOTAL = 0
    End If

    Call ofv.CERRARCONSULTA(RST2)




    'total encadenados
    strsql4 = "SELECT sum(A.Z) FROM "
    strsql4 = strsql4 & " ( "
    strsql4 = strsql4 & "SELECT SUM(1) AS Z FROM ORREGINFOBCRADET "
    strsql4 = strsql4 & " where secuencia > =  1 "
    strsql4 = strsql4 & " and  CODLOTE='" & LOTE & "'"
    strsql4 = strsql4 & "group by eventoraiz ) a"

    ' response.write strsql4
    Set rsT4 = ofv.crearconsultaEx(strsql4, cn, 1, parametros)
    If CLng("0" & rsT4(0)) > 0 Then
        totalencadenados = CLng("0" & rsT4(0))
    Else
        totalencadenados = TeventosEncadenados
    End If
    Call ofv.CERRARCONSULTA(rsT4)


    'total de agrupados
    strsql5="select sum(a.z) FROM ( SELECT  sum(1) as z  from  ORREGINFOBCRADET where  secuencia=-1 and CODLOTE='"&lote&"'  group by eventoraiz having sum(1) >1 ) a"
    Set rsT4 = ofv.crearconsultaEx(strsql5, cn, 1, parametros)

    If Not rsT4.EOF Then
        totalagr = CLng("0" & rsT4(0))
    Else
        totalagr = 0
    End If
    Call ofv.CERRARCONSULTA(rsT4)



    'total dif id repetitivos
    strsql4 = "select sum(z) from (SELECT sum(1) as z from  ORREGINFOBCRADET where secuencia=-1 and CODLOTE='"&lote&"' group by eventoraiz having sum(1) >1  ) a"
    ' response.write strsql4
    Set rsT4 = ofv.crearconsultaEx(strsql4, cn, 1, parametros)
    If Not rsT4.EOF Then
        totaldifrep = CLng("0" & rsT4(0))
    Else
        totaldifrep = 0
    End If
    Call ofv.CERRARCONSULTA(rsT4)



    totalrepetitivos = totaldifrep    '(tx-totalagr + totaldifrep) - totalagr

    'total dif id repetitivos
    strsql4 = "select sum(1) from (SELECT sum(1) as z from  ORREGINFOBCRADET where secuencia=-1 and CODLOTE='"&lote&"' group by eventoraiz having sum(1) >1  ) a"
    ' response.write strsql4
    Set rsT4 = ofv.crearconsultaEx(strsql4, cn, 1, parametros)
    If Not rsT4.EOF Then
        totaldifrep = CLng("0" & rsT4(0))
    Else
        totaldifrep = 0
    End If
    Call ofv.CERRARCONSULTA(rsT4)


    strsql4 = "SELECT SUM (TRCIMPBRUTO) FROM ORTRACER A "
    strsql4 = strsql4 & " WHERE A.OBJID IN ( "
    strsql4 = strsql4 & " SELECT DISTINCT NROEVENTO FROM "
    strsql4 = strsql4 & " ORRIERRLOTE "
    strsql4 = strsql4 & " WHERE CODLOTE='" & LOTE & "')"
    '    RESPONSE.WRITE STRSQL4
    Set rsT4 = ofv.crearconsultaEx(strsql4, cn, 1, parametros)
    If Not rsT4.EOF Then
        MONTOTOTALNOINF = rsT4(0)
    Else
        MONTOTOTALNOINF = 0
    End If
    Call ofv.CERRARCONSULTA(rsT4)






    STRSQL = "UPDATE ORREGINFOBCRA SET TOTAL =" & TOTAL
    STRSQL = STRSQL & ",totalMontoTri        =0" & Replace(TOTALMONTOTRI, ",", ".")
    STRSQL = STRSQL & ",TOTALINFORMAR        =0" & TX
    STRSQL = STRSQL & ",TOTALNOINFORMAR      =0" & TX2
    STRSQL = STRSQL & ",totrep               =0" & totalrepetitivos
    STRSQL = STRSQL & ",totenca              =0" & totalencadenados
    STRSQL = STRSQL & ",totdifrep            =0" & totaldifrep
    STRSQL = STRSQL & ",MONTOTOTALNOINF      =0" & REPLACE(MONTOTOTALNOINF&"", ",", ".")
    STRSQL = STRSQL & ",ESTADO               =2"
    STRSQL = STRSQL & ",ULTEVENTO            =" & ULTEVENTORAIZ & " "
    STRSQL = STRSQL & " WHERE CODLOTE='" & LOTE & "'"
    '''' RESPONSE.WRITE "<BR>" & STRSQL
    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)


    Call ofv.cerrarconn(cn)
    RESPONSE.REDIRECT "ORGENRUNOR.ASP?OP=PROCESADO"
    '

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FUNCIONES ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function EventosNormales(rs)

    POSI = 0

    Do While (Not rs.EOF)

        '  if rs(0)=1893 then

        Call VALIDALOTE2(rs, LOTE)
        '  end if
        '  posi = posi + 1
        rs.MOVENEXT

    Loop
    ' CALL GrabarReg(aRecordSave, ids, idspOS) 'se graban como repetitivos iguales dato pero no el monto


    If totalmontoRepet >= TOPEMONTO And REPESUMA = "S" Then

        Call GrabarReg(aRecordSave, ids, idspOS - 1, "1")                    'se graban como repetitivos iguales dato pero no el monto


    Else                  ' si no suman 1000 en el mes no van

        For iz = 0 To idspOS - 1
            STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
            STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"& ids(0, iz) &"','1002')"

            Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
        Next

    End If

    EventosNormales = fila
End Function



Function VALIDALOTE2(rs, LOTE)


    If rs("envia") = "S" Then


        XTipo = VerificarExtipo(rs, feciniT, fecfinT)
        ' if rs(0)= 1949 then response.write "<br><br><br><br>"&Xtipo&"<br><br><br><br>"
        ' XTipo N=Normal , R=Rectificado ,RH=Rectificado Hijo -- SEGUN SQL 3 CASOS DEL WHERE
        'RESPONSE.WRITE XTipo & "--"
        '     response.write rs("CARACTER") & "---" & Xtipo & "<br>"
        Select Case XTipo


            Case "N"    ' NORMAL


                REGVA = ValidaCampos(rs)


                '--------------------- TIPOS DE caracter -------------------------------------


                If (CStr(rs("CARACTER") & "") = "03" Or CStr(rs("CARACTER") & "") = "04") And REPESUMA = "S" Then    ' ----- eventos repetitivos

                    'if rs(0)= 1957 then response.write "<br><br><br><br>"&REGVA&"<br><br><br><br>"

                    If REGVA Then

                        If primerReg And CDbl(rs("TRCIMPBRUTO")) < TOPEMONTO Then

                            Call guardarRs(aRecordSave, rs, 0)
                            primerReg = False
                            fecha1 = CDate(rs("FECIDENT") + CLng(CDate("1-1-1900")))
                            mes = Month(fecha1)
                            CContable = Trim(rs("trcctaCble"))

                        End If

                        If CDbl(rs("TRCIMPBRUTO")) >= TOPEMONTO And REGVA Then

                            Set rShIJOSEnc = GetEncadenados(rs)
                            Call Grabar(rs, secuencia, rShIJOSEnc, strimestre, stANO)
                            Call ofv.CERRARCONSULTA(rShIJOSEnc)
                            '  CALL guardarRs(aRecordSave2,rs, 0)
                            '  idsex(0, 0) = CDBL(rs(0))
                            '  idsex(1, 0) = CDBL(REPLACE(RS("TRCIMPBRUTO"), ",", "."))
                            '  idspOSz=1
                            '  SET rShIJOSEnc = getEncadenados(rs)

                            '
                            '        REGVAENC = VALIDACAMPOS(rShIJOSEnc)


                            '          IF REGVAENC THEN
                            '      idsex(0, idspOSz) = CDBL(rShIJOSEnc(0))
                            '           idsex(1, idspOSz) = CDBL(rShIJOSEnc("TRCIMPBRUTO"))
                            '               idspOSz = idspOSz + 1
                            '
                            '               ELSE
                            '                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                            '                STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"&rShIJOSEnc(0) &"','1002')"
                            '
                            '               set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            '               END IF
                            '
                            '         rShIJOSEnc.movenext

                            ' loop
                            ' call ofv.cerrarconsulta(rShIJOSEnc)
                            ''
                            '  CALL GrabarReg(aRecordSave2, idsex, idspOSz-1)

                        Else

                            fecha2 = CDate(rs("FECIDENT") + CLng(CDate("1-1-1900")))
                            CContable2 = Trim(rs("trcctaCble"))

                            ' end if
                            If mes = Month(fecha2) And CContable = CContable2 Then    ' MIENTRAS ESTE EN EL MISMO MES LOS AGRUPO y ctacontable

                                If REGVA Then
                                    totalmontoRepet = totalmontoRepet + CDbl(rs("TRCIMPBRUTO"))

                                    ids(0, idspOS) = CDbl(rs(0))
                                    ids(1, idspOS) = CDbl(rs("TRCIMPBRUTO"))
                                    idspOS = idspOS + 1

                                    Set rShIJOSEnc = GetEncadenados(rs)

                                    Do While Not rShIJOSEnc.EOF

                                        REGVAENC = ValidaCampos(rShIJOSEnc)

                                        fecha2X = CDate(rShIJOSEnc("FECIDENT") + CLng(CDate("1-1-1900")))
                                        CContable2X = Trim(rShIJOSEnc("trcctaCble"))

                                        If REGVAENC Then

                                            If mes = Month(fecha2X) And CContable = CContable2X Then

                                                totalmontoRepet = totalmontoRepet + CDbl(rShIJOSEnc("TRCIMPBRUTO"))
                                                ids(0, idspOS) = CDbl(rShIJOSEnc(0))
                                                ids(1, idspOS) = CDbl(rShIJOSEnc("TRCIMPBRUTO"))
                                                idspOS = idspOS + 1

                                            Else
                                                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                                                STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"&rShIJOSEnc(0) &"','1002')"

                                                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                                            End If

                                        End If
                                        rShIJOSEnc.MOVENEXT

                                    Loop
                                    Call ofv.CERRARCONSULTA(rShIJOSEnc)
                                End If

                            Else    ' ME FIJO QUE SUME > TOPEMONTO cambio de mes
                                '  primerReg = TRUE



                                If totalmontoRepet >= TOPEMONTO Then

                                    Call GrabarReg(aRecordSave, ids, idspOS - 1, "2")    'se graban como repetitivos iguales dato pero no el monto


                                Else    ' si no suman 1000 en el mes no van

                                    For iz = 0 To idspOS - 1
                                        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                                        STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"& ids(0, iz) &"','1002')"

                                        Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                                    Next

                                End If

                                Call guardarRs(aRecordSave, rs, 0)
                                idspOS = 0

                                ids(0, idspOS) = CDbl(rs(0))
                                ids(1, idspOS) = CDbl(rs("TRCIMPBRUTO"))
                                idspOS = idspOS + 1
                                totalmontoRepet = CDbl(rs("TRCIMPBRUTO"))
                                fecha1 = CDate(rs("FECIDENT") + CLng(CDate("1-1-1900")))
                                mes = Month(fecha1)
                                CContable = Trim(rs("trcctaCble"))

                                Set rShIJOSEnc = GetEncadenados(rs)

                                Do While Not rShIJOSEnc.EOF

                                    REGVAENC = ValidaCampos(rShIJOSEnc)

                                    fecha2X = CDate(rShIJOSEnc("FECIDENT") + CLng(CDate("1-1-1900")))
                                    CContable2X = Trim(rShIJOSEnc("trcctaCble"))

                                    If REGVAENC Then

                                        If mes = Month(fecha2X) And CContable = CContable2X Then

                                            totalmontoRepet = totalmontoRepet + CDbl(rShIJOSEnc("TRCIMPBRUTO"))
                                            ids(0, idspOS) = CDbl(rShIJOSEnc(0))
                                            ids(1, idspOS) = CDbl(rShIJOSEnc("TRCIMPBRUTO"))
                                            idspOS = idspOS + 1

                                        Else
                                            STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                                            STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"&rShIJOSEnc(0) &"','1002')"

                                            Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                                        End If

                                    End If
                                    rShIJOSEnc.MOVENEXT

                                Loop
                                Call ofv.CERRARCONSULTA(rShIJOSEnc)





                            End If

                        End If
                    End If

                Else    ' ------eventos raiz

                    Set rShIJOS = GetEncadenados(rs)






                    If CDbl(rs("TRCIMPBRUTO")) < TOPEMONTO And rShIJOS.EOF and rs("TRCMARCASEND") <> "R" Then

                        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                        STRSQL = STRSQL & " VALUES ('" & LOTE & "','MONTO','" & rs(0) & "','1000')"
                        Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

                    ElseIf CDbl(rs("TRCIMPBRUTO")) >= TOPEMONTO or rs("TRCMARCASEND") = "R" Then

                        If REGVA Then

                            Call Grabar(rs, secuencia, rShIJOS, strimestre, stANO)         'como normal rShIJOS esta eof

                        End If


                    Else    ' CDBL(RS("TRCIMPBRUTO"))  < TOPEMONTO AND NOT rShIJOS.EOF

                        '   response.write ENCAMENORES
                        If ENCAMENORES = "S" Then
                            DTOTAL = CDbl(rs("TRCIMPBRUTO"))

                            If REGVA Then
                                Call Grabar(rs, secuencia, rShIJOS, strimestre, stANO)  'como normal
                            End If

                        Else

                            DTOTAL = CDbl(rs("TRCIMPBRUTO")) + CDbl(rs("trcimpotros")) + CDbl(rs("TRCIMPNOCTBLE")) + CDbl(rs("TRCIMPACRECEN"))


                            If DTOTAL >= TOPEMONTO Then

                                Set rShIJOSc = GetEncadenados(rs)
                                Do While Not rShIJOSc.EOF
                                    If rShIJOSc("envia") = "N" Then
                                        DTOTAL = DTOTAL - CDbl(rShIJOSc("TRCIMPBRUTO"))
                                    End If
                                    rShIJOSc.MOVENEXT
                                Loop
                                Call ofv.CERRARCONSULTA(rShIJOSc)
                            End If


                            '   if rs(0) = 1956 then response.write "<br>" & Dtotal  & CDBL(RS("TRCIMPBRUTO") )  + CDBL(RS("trcimpotros") )  + CDBL(RS("TRCIMPNOCTBLE") ) &" " & CDBL(RS("TRCIMPBRUTO") )  &"---"& CDBL(RS("trcimpotros") ) &"---"& CDBL(RS("TRCIMPNOCTBLE") )

                            If DTOTAL >= TOPEMONTO Then
                                If REGVA Then
                                    Call Grabar(rs, secuencia, rShIJOS, strimestre, stANO)         'como normal
                                End If

                            Else        ' NO LLEGA AL TOPEMONTO Y ENCAMENORES=N
                                '           if rs(0) = 1956 then response.write "<br>" & Dtotal  & CDBL(RS("TRCIMPBRUTO") )  + CDBL(RS("trcimpotros") )  + CDBL(RS("TRCIMPNOCTBLE") ) &" " & CDBL(RS("TRCIMPBRUTO") )  &"---"& CDBL(RS("trcimpotros") ) &"---"& CDBL(RS("TRCIMPNOCTBLE") )
                                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                                STRSQL = STRSQL & " VALUES ('" & LOTE & "','MONTO','" & rs(0) & "','1000')"
                                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                                '  SET rShIJOS2 = GetEncadenados(rs)
                                Call GrabarERRh(rShIJOS, 1000, "ENCADENADO")
                                ' CALL OFV.CERRARCONSULTA(rShIJOS2)

                            End If

                        End If

                    End If

                End If






        End Select

    Else   ' else de NO Envia


        If rs("trccodtracerp") & "" = "" Then  'SOLO SI ES RAIZ

            STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
            STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rs(0)&"','1001')"
            Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            Set rShIJOS2 = GetEncadenados(rs)
            '  IF RS(0)= 1 THEN RESPONSE.WRITE rShIJOS2.EOF
            Call GrabarERRh(rShIJOS2, 1001, "ESTADO DEL EVENTO RAIZ")
            Call ofv.CERRARCONSULTA(rShIJOS2)

        Else

            If rs("TRCMARCAACTP") = "N" Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rs(0)&"','1001')"
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If
        End If

    End If




End Function








Call ofv.CERRARCONSULTA(rs)
Call ofv.cerrarconn(cn)

End If
'-------------------------------------------------------------

Function GrabarReg(aRecordSave, ids, idspOS, WWWV)    'se graban como repetitivos iguales dato pero no el monto

    For iz = 0 To idspOS
'Select ORTRACER.*,E.descri as Sucursal,O.DESCRIpcion AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC

       ' CARACTER = aRecordSave(53, 0)
        CARACTER = aRecordSave(54, 0)
        FECORIG = aRecordSave(3, 0)
        FECIDENT = aRecordSave(4, 0)
        FECCONTA = aRecordSave(5, 0)
        FECFINAL = aRecordSave(23, 0)
        TRCOBSERV = Mid(aRecordSave(11, 0) & "", 1, 300)
      '  TIPOEVENTO = aRecordSave(55, 0)
         TIPOEVENTO = aRecordSave(56, 0)
      ' LINEGO2 = aRecordSave(52, 0)
        LINEGO2 = aRecordSave(53, 0)
       ' SUCURSAL = Mid(aRecordSave(49, 0) & "", 1, 50)
        SUCURSAL = Mid(aRecordSave(50, 0) & "", 1, 50)
       ' PROCESO = Mid(aRecordSave(50, 0) & "", 1, 50)
        PROCESO = Mid(aRecordSave(51, 0) & "", 1, 50)
       ' PRODSERV = aRecordSave(51, 0)
       PRODSERV = aRecordSave(52, 0)
        TRCSECSOPORTE = Mid(aRecordSave(27, 0) & "", 1, 50)
        TRCCTACBLE = Mid(aRecordSave(29, 0) & "", 1, 6)
        TRCIMPSEGURO = aRecordSave(21, 0)
       ' RIESGOASOC = aRecordSave(56, 0)
       RIESGOASOC = aRecordSave(57, 0)
        '    IF ids(0, iz)=1958 THEN
        ''response.write ids(0, iz) &"    " & TRCCTACBLE & " -" & WWWV & " <br> "
        STRSQL = "insert into ORREGINFOBCRADET "
        STRSQL = STRSQL & "(CODLOTE,EVENTORAIZ,SECUENCIA,CARACTER,FECINICIO,FECDESCUBRIMIENTO,FECREGCONTABLE,FECFINAL,DESCRIPCION,TIPOEVENTO,LINEANEGOCIO,SUCURSAL,PROCESO,PRODUCTO,MONTO,PARTIDACONTABLE,COBERPOLIZA,VINCOTROCRIESGO,AREASOPORTE,OBJIDORTRACER,TRIMESTRE,ESTADO,YANIO)"
        STRSQL = STRSQL & " values ("
        STRSQL = STRSQL & "'" & LOTE & "'," & EVENTORAIZ & ",'-1','" & CARACTER & "','" & FECORIG & "','" & FECIDENT & "',"
        STRSQL = STRSQL & "'" & FECCONTA & "','" & FECFINAL & "','" & TRCOBSERV & "','" & TIPOEVENTO & "',"
        STRSQL = STRSQL & "'" & LINEGO2 & "','" & SUCURSAL & "','" & PROCESO & "','" & PRODSERV & "','" & Replace(ids(1, iz), ",", ".") & "',"
        If CDbl("0" & TRCSECSOPORTE) > 0 Then
            AREASOP = 1
        Else
            AREASOP = 0
        End If
        strsql = strsql & "'" & TRCCTACBLE & "','"& TRCIMPSEGURO &"','"& RIESGOASOC&"','"&AREASOP &"',"
        STRSQL = STRSQL & "'" & ids(0, iz) & "','" & strimestre & "','N','" & stANO & "')"

       ' RESPONSE.WRITE STRSQL & " <BR>"

        Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
        POSI = POSI + 1
        TeventosRepetitivos = TeventosRepetitivos + 1
    Next
    ULTEVENTORAIZ = EVENTORAIZ
    EVENTORAIZ = EVENTORAIZ + 1
End Function



'-------------------------------------------------------------

Function VerificarExtipo(rs, feciniT, fecfinT)

    XVA = False

    If (rs("trccodtracerp") & "" = "") Then



        If FCARGA = "S" Then

            XVA = XVA Or ((rs("FECREGIS") >= feciniT And rs("FECREGIS") <= fecfinT))

        End If

        If FDESCUBRIMIENTO = "S" Then
            XVA = XVA Or ((rs("FECIDENT") >= feciniT And rs("FECIDENT") <= fecfinT))
        End If

        If FORIGEN = "S" Then
            XVA = XVA Or ((rs("FECORIG") >= feciniT And rs("FECORIG") <= fecfinT))
        End If

        If FCONTABLE = "S" Then

            XVA = XVA Or ((rs("FECCONTA") >= feciniT And rs("FECCONTA") <= fecfinT))

        End If


        If FFINAL = "S" Then

            XVA = XVA Or ((rs("FECFINENR") >= feciniT And rs("FECFINENR") <= fecfinT))

        End If






        If XVA Then
            VerificarExtipo = "N"    '& RS(0)

        Else
            ''   VerificarExtipo = "R" '& RS(0)
            VerificarExtipo = "N"    '& RS(0)
        End If
    Else
        VerificarExtipo = "R"    '& RS(0)
    End If



End Function

'------------------------------------------------------------------

Function guardarRs(Ar, Rse, Ix)

    For iz = 0 To Rse.fields.Count - 1
        ' response.write iz & "<br>"
        '  aRecordSave(iz, Ix) = CSTR( rse(iz) & "")
        Ar(iz, Ix) = CStr(Rse(iz) & "")
    Next

End Function



'------------------------------------------------------------------------------------------------

Function Grabar(rs, XEVENTORAIZ, rShIJOS, strimestre, stANO)    'como normal rShIJOS esta eof
    'Strsql = "Select *,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia From " & WKSETTABLA
    Dim xsecuencia

    If REGVA AND RS("TRCMARCASEND") <> "S" Then

        IF RS("TRCMARCASEND")="R" THEN

              STRSQL="SELECT A.EVENTORAIZ,A.SECUENCIA,A.TRIMESTRE,A.YANIO FROM ORREGINFOBCRADET A"
              STRSQL= STRSQL & " INNER JOIN ORTRACER B ON (A.OBJIDORTRACER=B.OBJID) "
              STRSQL=STRSQL & " WHERE B.CODTRACER='" & RS("TRCCODTRACERORG") & "'  ORDER BY A.OBJID DESC"
            '  response.write ""strsql
	          Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
	          IF NOT XXXC.EOF THEN
	           XEVENTORAIZ=XXXC("EVENTORAIZ")
	           XSECUENCIA=XXXC("SECUENCIA")
	           XTRIMESTRE=XXXC("TRIMESTRE")
	           XYANIO=XXXC("YANIO")
               XESTADO="R"
	          END IF
             CALL OFV.CERRARCONSULTA(XXXC)

        ELSE
         XEVENTORAIZ=EVENTORAIZ
	     XSECUENCIA=0
		 XTRIMESTRE=strimestre
	     XYANIO=stANO
	     XESTADO="N"
        END IF

        STRSQL = "insert into ORREGINFOBCRADET "
        STRSQL = STRSQL & "(CODLOTE,EVENTORAIZ,SECUENCIA,CARACTER,FECINICIO,FECDESCUBRIMIENTO,FECREGCONTABLE,FECFINAL,DESCRIPCION,TIPOEVENTO,LINEANEGOCIO,SUCURSAL,PROCESO,PRODUCTO,MONTO,PARTIDACONTABLE,COBERPOLIZA,VINCOTROCRIESGO,AREASOPORTE,OBJIDORTRACER,TRIMESTRE,ESTADO,YANIO)"
        STRSQL = STRSQL & " values ("
        STRSQL = STRSQL & "'" & LOTE & "'," & XEVENTORAIZ & ","&XSECUENCIA&",'" & rs("CARACTER") & "','" & rs("FECORIG") & "','" & rs("FECIDENT") & "',"
        STRSQL = STRSQL & "'" & rs("FECCONTA") & "','" & rs("FECFINAL") & "','" & Mid(rs("TRCOBSERV") & "", 1, 300) & "','" & rs("TIPOEVENTO") & "',"
        STRSQL = STRSQL & "'" & rs("LINEGO2") & "','" & Mid(rs("SUCURSAL") & "", 1, 50) & "','" & Mid(rs("PROCESO") & "", 1, 50) & "','" & Mid(rs("PRODSERV") & "", 1, 50) & "','" & Replace(rs("TRCIMPBRUTO"), ",", ".") & "',"

        If CDbl("0" & rs("TRCSECSOPORTE")) > 0 Then
            AREASOP = 1
        Else
            AREASOP = 0
        End If
        STRSQL = STRSQL & "'" & Mid(rs("TRCCTACBLE"), 1, 6) & "','" & rs("TRCIMPSEGURO") & "','" & rs("RIESGOASOC") & "','" & AREASOP & "',"
        STRSQL = STRSQL & "'" & rs(0) & "','" & Xtrimestre & "','"& XESTADO &"','" & XYANIO & "')"
        'if rs(0)=323 then response.write "<br>" & strsql & "<br>"
        Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

       IF RS("TRCMARCASEND") <> "R" THEN
        ULTEVENTORAIZ = EVENTORAIZ
        EVENTORAIZ = EVENTORAIZ + 1
        XEVENTORAIZ=EVENTORAIZ
       END IF
       xsecuencia=1
   else 'RAIZ YA MANDADOS POR X AJSUTES AGREGADOS O K NUEVOS ENCADENADOS

     if REGVA THEN



              STRSQL="SELECT EVENTORAIZ,SECUENCIA,TRIMESTRE,YANIO FROM ORREGINFOBCRADET WHERE OBJIDORTRACER=0" & RS(0) & " ORDER BY OBJID DESC"
	          Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
	          IF NOT XXXC.EOF THEN
	           XEVENTORAIZ=XXXC("EVENTORAIZ")
	           XSECUENCIA=XXXC("SECUENCIA")

	          STRSQL="SELECT MAX(SECUENCIA) FROM ORREGINFOBCRADET WHERE EVENTORAIZ=0" & XEVENTORAIZ
				  Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
				  IF NOT XXXC.EOF THEN

				   XSECUENCIA=XXXC(0) + 1
				  END IF



	          END IF
             CALL OFV.CERRARCONSULTA(XXXC)




     END IF

    End If
    If Not rShIJOS.EOF Then
        'xsecuencia = 1
        Do While Not rShIJOS.EOF

            regva2 = ValidaCampos(rShIJOS)

            If regva2 And rShIJOS("envia") = "S"   Then

               IF rShIJOS("TRCMARCASEND") <> "S" THEN

               IF rShIJOS("TRCMARCASEND")="J" THEN
                XESTADO="J"
               ELSE
               XESTADO="N"
               END IF
                STRSQL = "insert into ORREGINFOBCRADET "
                STRSQL = STRSQL & "(CODLOTE,EVENTORAIZ,SECUENCIA,CARACTER,FECINICIO,FECDESCUBRIMIENTO,FECREGCONTABLE,FECFINAL,DESCRIPCION,TIPOEVENTO,LINEANEGOCIO,SUCURSAL,PROCESO,PRODUCTO,MONTO,PARTIDACONTABLE,COBERPOLIZA,VINCOTROCRIESGO,AREASOPORTE,OBJIDORTRACER,TRIMESTRE,ESTADO,YANIO)"
                STRSQL = STRSQL & " values ("
                strsql = strsql & "'" & lote & "',"& XEVENTORAIZ &","&XSECUENCIA&",'"& RSHIJOS("CARACTER") &"','"& RSHIJOS("FECORIG")&"','"&RSHIJOS("FECIDENT")&"',"

                STRSQL = STRSQL & "'" & rShIJOS("FECCONTA") & "','" & rShIJOS("FECFINAL") & "','" & Mid(rShIJOS("TRCOBSERV") & "", 1, 300) & "','" & rShIJOS("TIPOEVENTO") & "',"
                STRSQL = STRSQL & "'" & rShIJOS("LINEGO2") & "','" & Mid(rShIJOS("SUCURSAL") & "", 1, 50) & "','" & Mid(rShIJOS("PROCESO") & "", 1, 50) & "','" & Mid(rShIJOS("PRODSERV") & "", 1, 50) & "','" & Replace(rShIJOS("TRCIMPBRUTO"), ",", ".") & "',"

                If CDbl("0" & rs("TRCSECSOPORTE")) > 0 Then
                    AREASOP = 1
                Else
                    AREASOP = 0
                End If
                STRSQL = STRSQL & "'" & Mid(rShIJOS("TRCCTACBLE") & "", 1, 6) & "','" & rShIJOS("TRCIMPSEGURO") & "','" & rShIJOS("RIESGOASOC") & "','" & AREASOP & "',"
                STRSQL = STRSQL & "'" & rShIJOS(0) & "','" & strimestre & "','"& XESTADO &"','" & stANO & "')"
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                xsecuencia = xsecuencia + 1
                POSI = POSI + 1
                TeventosEncadenados = TeventosEncadenados + 1
              ELSE

              END IF

            Else


                If rShIJOS("envia") <> "S" Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rshijos(0)&"','1001')"
                End If
                REGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

            End If
            rShIJOS.MOVENEXT


        Loop



    End If
    'ULTEVENTORAIZ = EVENTORAIZ
    'EVENTORAIZ = EVENTORAIZ + 1
End Function

'-----------------------------------------------------------------

Function GrabarERRh(rShIJOSZ, ERROR, estado)    'como normal rShIJOS esta eof
    'Strsql = "Select *,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia From " & WKSETTABLA

    ' IF RS(0)=1956 THEN RESPONSE.WRITE RSHIJOSZ.EOF & "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
    Do While Not rShIJOSZ.EOF

        '     IF rshijosZ(0)=1898 THEN
        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
        STRSQL = STRSQL & " VALUES ('"&LOTE&"','"&estado&"','"&rshijosZ(0)&"','"& ERROR &"')"
        REGVA = False
        Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

        If rShIJOSZ("ENVIA") <> "S" Then
            STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
            STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rshijosZ(0)&"','1001')"
            REGVA = False
            Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
        End If



        '   RESPONSE.WRITE STRSQL & "<BR>"
        '   END IF
        xxx = ValidaCamposH(rShIJOSZ)
        rShIJOSZ.MOVENEXT


    Loop



End Function


'----------------------------------------------------------------

Function ValidaCampos(rs)
    Dim BREGVA

    BREGVA = True



    Select Case rs("CARACTER")


        Case "05"

            If rs("FECORIG") > 0 Then

                If rs("FECIDENT") < rs("FECORIG") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If


            If rs("FECFINAL") > 0 Then
                If rs("FECFINAL") < rs("FECIDENT") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
                    '    RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If




            If Not IsNumeric(rs("CARACTER")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If



            If Trim(rs("trcctaCble") & "") = "" Then

                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','10')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)


            End If



        Case Else
            If rs("FECORIG") > 0 Then

                If rs("FECIDENT") < rs("FECORIG") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If

            If rs("FECFINAL") > 0 Then
                If rs("FECFINAL") < rs("FECIDENT") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
                    '    RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If



            If Not IsNumeric(rs("TRCIMPSEGURO")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','COBERTURA DE SEGURO','"&rs(0)&"','5')"
                '    RESPONSE.WRITE STRSQL
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                BREGVA = False
            End If

            If False Then
                ' IF BUSCAEXCAR(RS("trcobserv"), "") THEN
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','DESCRIPCION','"&rs(0)&"','10')"
                '    RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If

            If rs("SUCURSAL") & "" = "" Then
                'IF BUSCAEXCAR(RS("SUCURSAL"), "") OR RS("SUCURSAL") & "" = ""  THEN
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','SUCURSAL','"&rs(0)&"','10')"
                '    RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If

            If rs("PRODSERV") & "" = "" Then
                ' IF BUSCAEXCAR(RS("PRODSERV"), "") OR RS("PRODSERV") &""= "" THEN
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','PRODUCTO','"&rs(0)&"','10')"
                '    RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If

            If rs("PROCESO") & "" = "" Then
                '  IF BUSCAEXCAR(RS("PROCESO"), "") OR RS("PROCESO")& "" =""  THEN
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','PROCESO','"&rs(0)&"','10')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If



            If Not IsNumeric(rs("LINEGO2")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','LINEA DE NEGOCIO','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If


            If Not IsNumeric(rs("CARACTER")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If



            If Not IsNumeric(rs("TIPOEVENTO")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','TIPO DE EVENTO','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If

            If rs("CARACTER") = "01" Or rs("CARACTER") = "03" Or rs("CARACTER") = "05" Or rs("CARACTER") = "06" Or rs("CARACTER") = "07" Then

                If Trim(rs("trcctaCble") & "") = "" Then

                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

                End If
            End If


            If rs("CARACTER") = "02" Or rs("CARACTER") = "04" Or rs("CARACTER") = "08" Or rs("CARACTER") = "09" Then

                If Trim(rs("trcctaCble") & "") <> "" Then

                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

                End If


                If rs("FECCONTA") > 0 Then

                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA CONTABLE','"&rs(0)&"','18')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

                End If

            End If

            '  IF NOT ISNUMERIC (rs("RIESGOASOC")) THEN

            '        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
            ''        STRSQL = STRSQL & " VALUES ('"&LOTE&"','RIESGO ASOCIADO','"&rs(0)&"','09')"
            '        '     RESPONSE.WRITE STRSQL
            '        BREGVA = FALSE
            '        set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
            '  END IF

    End Select

    ValidaCampos = BREGVA

    If BREGVA = False Then
        Set rShIJOS3 = GetEncadenados(rs)
        Call GrabarERRh(rShIJOS3, 1003, "ENCADENADO")
        Call ofv.CERRARCONSULTA(rShIJOS3)
    End If

End Function

Function ValidaCamposH(rs)
    Dim BREGVA


    BREGVA = True


    Select Case rs("CARACTER")


        Case "05"

            If rs("FECORIG") > 0 Then

                If rs("FECIDENT") < rs("FECORIG") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If


            If rs("FECFINAL") > 0 Then
                If rs("FECFINAL") < rs("FECIDENT") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
                    '    RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If




            If Not IsNumeric(rs("CARACTER")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If



            If Trim(rs("trcctaCble") & "") = "" Then

                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','10')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)


            End If



        Case Else

            If rs("FECORIG") > 0 Then

                If rs("FECIDENT") < rs("FECORIG") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If

            If rs("FECFINAL") > 0 Then
                If rs("FECFINAL") < rs("FECIDENT") Then
                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
                    '    RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                End If
            End If


            If Not IsNumeric(rs("TRCIMPSEGURO")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','COBERTURA DE SEGURO','"&rs(0)&"','5')"
                '    RESPONSE.WRITE STRSQL
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
                BREGVA = False
            End If


            If False Then
                '   IF BUSCAEXCAR(RS("trcobserv"), "") THEN



                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','DESCRIPCION','"&rs(0)&"','10')"
                '    RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If

            If rs("SUCURSAL") & "" = "" Then
                ' IF BUSCAEXCAR(RS("SUCURSAL"), "") OR RS("SUCURSAL") & "" = ""  THEN
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','SUCURSAL','"&rs(0)&"','10')"
                '    RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If

            If rs("PRODSERV") & "" = "" Then
                'IF BUSCAEXCAR(RS("PRODSERV"), "") OR RS("PRODSERV") &""= "" THEN
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','PRODUCTO','"&rs(0)&"','10')"
                '    RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If

            If rs("PROCESO") & "" = "" Then
                'IF BUSCAEXCAR(RS("PROCESO"), "") OR RS("PROCESO")& "" =""  THEN
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','PROCESO','"&rs(0)&"','10')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If



            If Not IsNumeric(rs("LINEGO2")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','LINEA DE NEGOCIO','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If


            If Not IsNumeric(rs("CARACTER")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If



            If Not IsNumeric(rs("TIPOEVENTO")) Then
                STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                STRSQL = STRSQL & " VALUES ('"&LOTE&"','TIPO DE EVENTO','"&rs(0)&"','09')"
                '     RESPONSE.WRITE STRSQL
                BREGVA = False
                Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)
            End If


            If rs("CARACTER") = "01" Or rs("CARACTER") = "03" Or rs("CARACTER") = "05" Or rs("CARACTER") = "06" Or rs("CARACTER") = "07" Then

                If Trim(rs("trcctaCble") & "") = "" Then

                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

                End If
            End If

            If rs("CARACTER") = "02" Or rs("CARACTER") = "04" Or rs("CARACTER") = "08" Or rs("CARACTER") = "09" Then

                If Trim(rs("trcctaCble") & "") <> "" Then

                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

                End If


                If rs("FECCONTA") > 0 Then

                    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                    STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA CONTABLE','"&rs(0)&"','18')"
                    '     RESPONSE.WRITE STRSQL
                    BREGVA = False
                    Set XXXC = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

                End If

            End If

    End Select
    ValidaCamposH = BREGVA
End Function


Function GetEncadenados(rs)
    If (SESSION("APLMDLOR") & "") = "BE" Then
        STRSQL = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRIpcion AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
    Else
        STRSQL = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
    End If


    'strsql= strsql & " inner join orestados on (ORTRACER.trcestado=orestados.codigo and orestados.envia='S')"
    STRSQL = STRSQL & " inner join orestados on (ORTRACER.trcestado=orestados.codigo )"
    STRSQL = STRSQL & " left join eco E on (ORTRACER.trcecoresp=E.eco)  "


    If (SESSION("APLMDLOR") & "") = "BE" Then
        STRSQL = STRSQL & " LEFT JOIN TABPROCESOS01 O ON (ORTRACER.trcPROCESO=O.OBJID)"
    Else
        STRSQL = STRSQL & " LEFT JOIN OBJETOS O ON (ORTRACER.trcPROCESO=O.CODOBJETO)"
    End If


    STRSQL = STRSQL & " left join ORTIPOPRODUCTO OT on (ORTRACER.trcPRODUCTO=OT.OBJID)  "
    STRSQL = STRSQL & " left join ORTIPOCATEGO OTCAT on (ORTRACER.TRCEVENTORIESGO2 = OTCAT.OBJID)  "
    STRSQL = STRSQL & " left join ORRIESGOASOC ORAS on (ORTRACER.TRCRIESGOASOC = ORAS.OBJID)  "

    STRSQL = STRSQL & " left join ORLINNEGO2 OL2 on (ORTRACER.trcNEGOLINE=OL2.OBJID)  "
    STRSQL = STRSQL & " left join ORTIPOCARACTER OC on (ORTRACER.trcCARACTER=OC.OBJID)  "
    STRSQL = STRSQL & " WHERE ORTRACER.TRCCODTRACERP = '" & rs("CODTRACER") & "' AND ("

    'IF FCARGA = "S" AND FDESCUBRIMIENTO = "S" THEN
    ' STRSQL = STRSQL & SORFECINIHW2 &" OR " & SORFECINIHW3
    ' ELSEIF FDESCUBRIMIENTO = "S" THEN
    '       STRSQL = STRSQL & SORFECINIHW3
    ' ELSE
    '       STRSQL = STRSQL & SORFECINIHW2
    ' END IF

    STRSQL = STRSQL & SESSION(WSPSESION)    '& SORFECINIHW1

    SORFECINIHW = " (TRCFECACTP > = " & feciniT & " and TRCFECACTP < = " & fecfinT & ") "
    STRSQL = STRSQL & "  OR " & SORFECINIHW
    STRSQL = STRSQL & ") ORDER BY ORTRACER.FECIDENT,ORTRACER.FECREGIS,oc.cODBCRA,ORTRACER.TRCCTACBLE"
  '  response.write "<BR>QQQ" & STRSQL & "QQ<BR>"
    Set GetEncadenados = ofv.crearconsultaEx(STRSQL, cn, 1, parametros)

End Function


Function SqlEventosNormales()



    enter=chr(13)

    If (SESSION("APLMDLOR") & "") = "BE" Then
        STRSQL = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRIpcion AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
    Else
        STRSQL = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
    End If

    'strsql= strsql & " inner join orestados on (ORTRACER.trcestado=orestados.codigo and orestados.envia='S')"
    STRSQL = STRSQL & " inner join orestados on (ORTRACER.trcestado=orestados.codigo )" & enter
    STRSQL = STRSQL & " left join eco E on (ORTRACER.trcecoresp=E.eco)  " & enter

    If (SESSION("APLMDLOR") & "") = "BE" Then
        STRSQL = STRSQL & " LEFT JOIN TABPROCESOS01 O ON (ORTRACER.trcPROCESO=O.OBJID)"& enter
    Else
        STRSQL = STRSQL & " LEFT JOIN OBJETOS O ON (ORTRACER.trcPROCESO=O.CODOBJETO)"& enter
    End If


    STRSQL = STRSQL & " left join ORTIPOPRODUCTO OT on (ORTRACER.trcPRODUCTO=OT.OBJID)  "& enter
    STRSQL = STRSQL & " left join ORTIPOCATEGO OTCAT on (ORTRACER.TRCEVENTORIESGO2 = OTCAT.OBJID)  "& enter
    STRSQL = STRSQL & " left join ORRIESGOASOC ORAS on (ORTRACER.TRCRIESGOASOC = ORAS.OBJID)  "& enter

    STRSQL = STRSQL & " left join ORLINNEGO2 OL2 on (ORTRACER.trcNEGOLINE=OL2.OBJID)  "& enter
    STRSQL = STRSQL & " left join ORTIPOCARACTER OC on (ORTRACER.trcCARACTER=OC.OBJID)  "& enter

    If FCARGA <> "N" Then



        SORFECINIHW = " (ORTRACER.FECREGIS > = " & feciniT & " and ORTRACER.FECREGIS < = " & fecfinT & ") "

        SORFECINIHW1 = " (b.FECREGIS > = " & feciniT & " and b.FECREGIS < = " & fecfinT & ")" & SCONECTOR

        SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
        SCONECTOR = " OR "

    End If


    If FDESCUBRIMIENTO <> "N" Then



        SORFECINIHW = " (ORTRACER.FECIDENT > = " & feciniT & " and ORTRACER.FECIDENT < = " & fecfinT & ") "
        SORFECINIHW1 = SORFECINIHW1 & SCONECTOR & " (b.FECIDENT > = " & feciniT & " and b.FECIDENT < = " & fecfinT & ") "

        SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
        SCONECTOR = " OR "

    End If



    If FORIGEN <> "N" Then



        SORFECINIHW = " (ORTRACER.FECORIG > = " & feciniT & " and ORTRACER.FECORIG < = " & fecfinT & ") "
        SORFECINIHW1 = SORFECINIHW1 & SCONECTOR & " (b.FECORIG > = " & feciniT & " and b.FECORIG < = " & fecfinT & ") "
        SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
        SCONECTOR = " OR "

    End If

    If FCONTABLE <> "N" Then



        SORFECINIHW = " (ORTRACER.FECCONTA > = " & feciniT & " and ORTRACER.FECCONTA < = " & fecfinT & ") "
        SORFECINIHW1 = SORFECINIHW1 & SCONECTOR & " (b.FECCONTA > = " & feciniT & " and b.FECCONTA < = " & fecfinT & ") "
        SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
        SCONECTOR = " OR "

    End If


    If FFINAL <> "N" Then



        SORFECINIHW = " (ORTRACER.FECFINENR > = " & feciniT & " and ORTRACER.FECFINENR < = " & fecfinT & ") "
        SORFECINIHW1 = SORFECINIHW1 & SCONECTOR & " (b.FECFINENR > = " & feciniT & " and b.FECFINENR < = " & fecfinT & ") "
        SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
        SCONECTOR = " OR "

    End If


    swhere = "  where NOT (ortracer.codTRACER like 'DOBLIGA%') "
    swhere = swhere & "  AND (ortracer.trccodtracerp IS NULL OR  ortracer.trccodtracerp = '') "



    If SESSION(WSPSESION) <> "" Then
      '  swhere = swhere & "  and ( ortracer.TRCMARCASEND <> 'S' AND  (" & SESSION(WSPSESION) & ") ) "
         swhere = swhere & "  and ( ortracer.TRCMARCASEND not in ('S')  AND  (" & SESSION(WSPSESION) & ") ) "

    End If


    STRSQL = STRSQL & swhere

    STRSQL = STRSQL & "  or (ortracer.TRCMARCAACTP in ('K','X')  AND (ORTRACER.TRCFECACTP > = " & feciniT & " and ORTRACER.TRCFECACTP < = " & fecfinT & "))"


    '  STRSQL = STRSQL & "  or (ortracer.TRCMARCAACTP <> 'S' AND (ORTRACER.TRCFECACTP > = " & feciniT & " and ORTRACER.TRCFECACTP < = " & fecfinT & "))"


    '  STRSQL = STRSQL & "  or (ortracer.TRCMARCAACTH <> 'S' AND (ORTRACER.TRCFECACTH > = " & feciniT & " and ORTRACER.TRCFECACTH < = " & fecfinT & "))"
    '  STRSQL = STRSQL & "  or (ortracer.TRCMARCAACTH <> 'S' AND (ORTRACER.TRCFECACTH > = " & feciniT & " and ORTRACER.TRCFECACTH < = " & fecfinT & "))"

    STRSQL = STRSQL & " ORDER BY ORTRACER.TRCCTACBLE,oc.CODBCRA,ORTRACER.FECIDENT,ORTRACER.FECREGIS "

    SqlEventosNormales = STRSQL


End Function

%>
