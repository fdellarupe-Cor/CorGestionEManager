<%

nFecini = 0
xstyle = " style='font-size: 10pt; font-family: Verdana;' "
If nFecini < CLng(CDate("01-01-1900")) Then
   nFecini = CLng(CDate("01-01-1900"))
Else
   nFecini = 0
End If

Function ORRECALCEVENTO(XORCODIGO, XORMODAL)

    Dim XORTABCARACTER(100, 2)
    XORTABIND = 0
    XORTOPIND = 0

    XORSENENCA = "NO"

    XORIMPOM = 0
    XORIMPRE = 0
    XORIMPPO = 0
    XORIMPDE = 0
    XORIMPNC = 0

    XORCODCARACTER = 0
    XOROPCNETO = "S"
    XORESHIJO = "N"
    Simpcont = "X"
    StrSql = "select TRCCODTRACERP,TRCCARACTER FROM ORTRACER  WHERE CODTRACER = '" & XORCODIGO & "'    "
    ''  response.write strsql & "<br>"
    Set Xrs2 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not Xrs2.EOF Then

        XORCODCARACTER = Xrs2(1) & ""

        If (Xrs2(0) & "") <> "" Then

            XORSENENCA = "SI"
            XORCODIGORAIZ = Xrs2(0)
            XORESHIJO = "S"
            sFECREGIS = CLng(Date) - CLng(CDate("01-01-1900"))

            StrSql = "UPDATE ORTRACER SET  "
            StrSql = StrSql & " TRCMARCAACTH = 'A', TRCFECACTH = '" & sFECREGIS & "' "
            StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "


            ''  RESPONSE.WRITE Strsql & "<BR>"
            ''''     call ofv.crearconsultaEx(StrSql,cn,1,volver)

            ''    XORESHIJOxx     = 0
            ''   sFECREGIS     = sFECREGIS  / XORESHIJOxx

        End If




        StrSql = "Select * From ORTIPOCARACTER WHERE OBJID = 0" & XORCODCARACTER
        ''response.write strsql & chr(13)
        Set Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not Xrs.EOF Then
            Simpcont = Xrs("IMPCONT") ' ME FIJO SI TIENE IMPACTO CONTABLE
            '''''        'response.write "PASO2<BR>"
            If (Xrs(2) & "") = "NCN" Then XOROPCNETO = "N"
        End If


    End If

    If XORMODAL = "SI" Then

        XORSENENCA = "SI"
        XORCODIGORAIZ = XORCODIGO

    End If


    If XORSENENCA = "SI" Then


        StrSql = "Select * From ORTIPOCARACTER ORDER BY OBJID"
        Set Xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not Xrs.EOF Then

            Do While Not Xrs.EOF
                XORTABIND = XORTABIND + 1

                XORTABCARACTER(XORTABIND, 0) = Xrs(0) 'CODIGO
                XORTABCARACTER(XORTABIND, 1) = Xrs(2) 'ACUMULA


                XORTOPIND = XORTABIND

                Xrs.movenext
            Loop


        End If

        ''' CARACTER   --- 24


        ''' MON.BRUTO  --- 12

        ''' OTR.MONTOS --- 14
        ''' RECUPEROS  --- 15
        ''' PERDIDAS   --- 13
        ''' DESAFECTA  --- 36
        ''' NO CONTABLES -- 42 - TRCIMPNOCTBLE - NCN


        StrSql = "select trcimpbruto,trccaracter FROM ORTRACER  WHERE TRCCODTRACERP = '" & XORCODIGORAIZ & "' AND TRCESTADO <> 'CAN'  "
        ''  'response.write strsql & "<br>"
        Set Xrs2 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not Xrs2.EOF Then

            Do While Not Xrs2.EOF


                XORWKCARACTER = CLng("0" & Xrs2(1))
                XORWKimpbruto = CDbl(Xrs2(0))

                XORWKCARACMON = "OM"


                For XORTABIND = 1 To XORTOPIND

                    If XORTABCARACTER(XORTABIND, 0) = XORWKCARACTER Then
                        XORWKCARACMON = XORTABCARACTER(XORTABIND, 1)
                        Exit For
                    End If

                Next

                ''  XORIMPOM = 0
                ''  XORIMPRE = 0
                ''  XORIMPPO = 0
                ''  XORIMPDE = 0
                ''  XORIMPNC = 0


                Select Case XORWKCARACMON
                    Case "OM"

                        XORIMPOM = XORIMPOM + XORWKimpbruto

                    Case "REC"

                        XORIMPRE = XORIMPRE + XORWKimpbruto

                    Case "DES"

                        XORIMPDE = XORIMPDE + XORWKimpbruto

                    Case "PER"

                        XORIMPPO = XORIMPPO + XORWKimpbruto

                    Case "NCN"

                        XORIMPNC = XORIMPNC + XORWKimpbruto

                    Case Else

                        XORIMPOM = XORIMPOM + XORWKimpbruto

                End Select

                Xrs2.movenext

            Loop

        End If



        StrSql = "UPDATE ORTRACER SET "
        StrSql = StrSql & " TRCIMPOTROS = '" & Replace(XORIMPOM & "", ",", ".") & "' , "
        StrSql = StrSql & " trcimpacrecen = '" & Replace(XORIMPPO & "", ",", ".") & "' , "
        StrSql = StrSql & " TRCIMPDESAFECTA = '" & Replace(XORIMPDE & "", ",", ".") & "' , "
        StrSql = StrSql & " TRCIMPRECUPERA = '" & Replace(XORIMPRE & "", ",", ".") & "', "
        StrSql = StrSql & " TRCIMPNOCTBLE = '" & Replace(XORIMPNC & "", ",", ".") & "' "
        StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
        'response.write "11 " & strsql & chr(13)
        Call ofv.crearconsultaEx(StrSql, cn, 1, volver)

        If XOROPCNETO = "N" Then
            '    if Simpcont="N"   then
            '     Strsql = "UPDATE ORTRACER SET TRCIMPNETO = "
            '        Strsql = Strsql & " (  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
            '     Strsql = Strsql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            '    else
            '    Strsql = "UPDATE ORTRACER SET TRCIMPNETO = "
            '    Strsql = Strsql & " ( TRCIMPBRUTO +  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
            '    Strsql = Strsql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            '    end if
            '    response.write "1 " & strsql & chr(13) & Simpcont & chr(13)
            '    call ofv.crearconsultaEx(StrSql,cn,1,volver)


            If XORESHIJO = "S" Then

                StrSql = "UPDATE ORTRACER SET  "
                StrSql = StrSql & " TRCIMPNOCTBLE = (  TRCIMPNOCTBLE  + " & Replace(XORIMPNC & "", ",", ".") & " ) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "

                '    RESPONSE.WRITE Strsql & "<BR>"
                ''    call ofv.crearconsultaEx(StrSql,cn,1,volver)



                '     response.write "2 " & strsql & chr(13)


            Else

                ' if Simpcont="N" then
                ' Strsql = "UPDATE ORTRACER SET  "
                ' Strsql = Strsql & " TRCIMPNOCTBLE = (  TRCIMPNOCTBLE   ) "
                ' Strsql = Strsql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
                ' else
                '   Strsql = "UPDATE ORTRACER SET  "
                '    Strsql = Strsql & " TRCIMPNOCTBLE = (  TRCIMPNOCTBLE  + TRCIMPBRUTO ) "
                '    Strsql = Strsql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
                ' end if
                ''    RESPONSE.WRITE Strsql & "<BR>"
                ' call ofv.crearconsultaEx(StrSql,cn,1,volver)

                'response.write "3 " & strsql & chr(13)
                '     STRSQL = "select B.IMPCONT from ortracer A "
                '          STRSQL=STRSQL & " INNER JOIN ortipocaracter B ON(A.TRCCARACTER=B.OBJID)"
                '          Strsql = Strsql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
                '                   response.write "12" & strsql & chr(13)
                '          set Xrs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                '          IF NOT Xrs2.eof then
                '               Simpcont=Xrs2(0)
                '          else
                '               Simpcont="S"
                '     end if

            End If

            StrSql = "select B.IMPCONT from ortracer A "
            StrSql = StrSql & " INNER JOIN ortipocaracter B ON(A.TRCCARACTER=B.OBJID)"
            StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            'response.write "12" & strsql & chr(13)
            Set Xrs2 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
            If Not Xrs2.EOF Then
                Simpcont = Xrs2(0)
            Else
                Simpcont = "S"
            End If

            If Simpcont = "N" Then
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " (  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            Else
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " ( TRCIMPBRUTO +  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            End If
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)

            'response.write "1 " & strsql & chr(13) & Simpcont & chr(13)




        Else 'XORSENENCA = "NO"


            StrSql = "select B.IMPCONT from ortracer A "
            StrSql = StrSql & " INNER JOIN ortipocaracter B ON(A.TRCCARACTER=B.OBJID)"
            StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            'response.write "12" & strsql & chr(13)
            Set Xrs2 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
            If Not Xrs2.EOF Then
                Simpcont = Xrs2(0)
            Else
                Simpcont = "S"
            End If


            If Simpcont = "N" Then
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " (  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            Else
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " ( TRCIMPBRUTO + TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGORAIZ & "' "
            End If


            'response.write "4 " & strsql & chr(13)
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)


        End If

    Else

        If XOROPCNETO = "N" Then

            If Simpcont = "N" Then
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " (  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGO & "' "
            Else
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " (  TRCIMPBRUTO + TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGO & "' "
            End If
            'response.write "5 " & strsql & chr(13)
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)

        Else
            If Simpcont = "N" Then
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " (  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGO & "' "
            Else
                StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
                StrSql = StrSql & " ( TRCIMPBRUTO +  TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
                StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGO & "' "
            End If
            'response.write "6 " & strsql & chr(13)
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)

        End If



    End If

    If XORMODAL = "NO" Then

        If XOROPCNETO = "S" Then
            '''''       RESPONSE.WRITE "PASO4<BR>"

            StrSql = "UPDATE ORTRACER SET TRCIMPNETO = "
            StrSql = StrSql & " ( TRCIMPBRUTO + TRCIMPOTROS - TRCIMPRECUPERA + trcimpacrecen - TRCIMPDESAFECTA) "
            StrSql = StrSql & " WHERE  CODTRACER = '" & XORCODIGO & "' "

            'response.write "9 " & strsql & chr(13)
            Call ofv.crearconsultaEx(StrSql, cn, 1, volver)

        End If

    End If

    'response.write "XOROPCNETO="&XOROPCNETO &"       Simpcont= " & Simpcont & "    XORSENENCA=" & XORSENENCA & "      XORESHIJO=" & XORESHIJO & "           XORMODAL=" & XORMODAL & "   XORCODIGO=" & XORCODIGO
End Function




Function getobjurl(getnCodigo, getstipo)


    WKURLCONS = getnCodigo & "-" & getstipo

    Set getcn = ofv.conectar(ofv.strconn2)


    xStrSql = "SELECT * FROM TIPOBJETO  "
    xStrSql = xStrSql & " Where  tipobjeto = '" & getstipo & "' "

    Set getrs = ofv.crearconsultaEx(xStrSql, getcn, 1, parametros)
    If Not getrs.EOF Then

        wkhosturl = request.servervariables("HTTP_REFERER")

        wkhosturlX = UCase(wkhosturl & "NO")

        wkhosturln = InStr(wkhosturlX, "/COREMANAGER/")

        If wkhosturln > 0 Then

            wkhosturl = Left(wkhosturl, wkhosturln)


        Else

            wkhosturln = InStr(UCase(request.servervariables("SCRIPT_NAME") & "NO"), "/COREMANAGER/")

            wkhosturl = Left(request.servervariables("SCRIPT_NAME") & "NO", wkhosturln)

            wkhosturl = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & wkhosturl


        End If

        WKURLCONS = wkhosturl


        WKURLCONS = WKURLCONS & "coress20/corexpress.asp?opt=CC-" & getnCodigo & "-" & getstipo

    End If

    getobjurl = WKURLCONS


End Function


Function EsqnomP(xtipo, xcodigo, ofv)

    Set xcn = ofv.conectar(ofv.strconn2)


    If xtipo = "RSK" Then

        StrSql = "SELECT * FROM  ORRSKCIRCUITO "
        StrSql = StrSql & " Where OBJID = 0" & xcodigo
        Set Xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
        If Not Xrs.EOF Then
            If Xrs("APLIC") = "OR" Then
                xdescri = "CIRCUITO AUTOEVALUACION"
            Else
                xdescri = "PLAN DE CONTINGENCIAS"
            End If

            StrSql = "Select * From objetos where codobjeto = 0" & Xrs("CODOBJETO")
            Set Xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
            If Not Xrs.EOF Then
                xiden = Xrs(1)
                xobjdes = Xrs(3)
            End If


        End If
        Call ofv.cerrarconsulta(Xrs)

    Else




        StrSql = "Select descri From tipobjeto where tipobjeto ='" & xtipo & "' "
        Set Xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
        If Not Xrs.EOF Then xdescri = Xrs(0)
        Call ofv.cerrarconsulta(Xrs)

        StrSql = "Select * From objetos where codobjeto =" & xcodigo
        ' response.write strsql
        Set Xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
        If Not Xrs.EOF Then
            xiden = Xrs(1)
            xobjdes = Xrs(3)
        End If
        Call ofv.cerrarconsulta(Xrs)

    End If





    Call ofv.cerrarconn(xcn)

    EsqnomP = xdescri & ": " & xobjdes

End Function


Function msgsend(xusr, xmsgusr, ofv, XNCODIGO, XSTIPO, xsnovcodi)


    xmsgusrn = Replace(xmsgusr, ">>>", "<br>")
    xmsgusr = Replace(xmsgusr, ">>>", " ")

    Set xcn = ofv.conectar(ofv.strconn0)

    StrSql = "Select ID, descripcion From usuarios where ID ='" & xusr & "' "
    Set Xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)

    If Xrs.EOF Then
        Call ofv.cerrarconsulta(Xrs)
    Else

        Xdestinodesc = Xrs(1)
        Call ofv.cerrarconsulta(Xrs)
    End If

    Call ofv.cerrarconn(xcn)

    Set xcn = ofv.conectar(ofv.strconn3)

    YC = 1
    YD = 0
    Xmarca = 0

    StrSql = " select * from CsMsgSearch where Msgusrid ='" & xusr & "'  order by msgid DESC"
    Set Xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
    Do Until Xrs.EOF

        If YC > 14 Then
            ReDim Preserve Xmsgid(YD)
            Xmsgid(YD) = Xrs(0)
            YD = YD + 1
            Xmarca = 1
        End If

        Xrs.movenext
        YC = YC + 1

    Loop
    Call ofv.cerrarconsulta(Xrs)

    If Xmarca = 1 Then
        For YC = 0 To UBound(Xmsgid)
            StrSql = " delete  from MsgChat where Msgid =" & Xmsgid(YC)
            Call ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
        Next
    End If

    StrSql = "insert into MsgChat (Msgusrid,Msgusrdesc,Msgtstamp,Msgdata,Msgusrorigid,Msgusroridesc) Values('"
    StrSql = StrSql & xusr & "','" & Xdestinodesc & "','" & Now & "','" & xmsgusr & "','" & session("usuario") & "','" & session("segusuario") & "') "
    Call ofv.crearconsultaEx(StrSql, xcn, 1, parametros)

    Call ofv.cerrarconn(xcn)

    Set xcn = ofv.conectar(ofv.strconn0)

    StrSql = "Select email From Usuarios  where ID = 'COREMANAGER' "
    Set rsx1 = ofv.crearconsultaEx(StrSql, xcn, 1, volver)
    If Not rsx1.EOF Then
        If Len(rsx1(0)) > 0 Then
            xfromusr = "COREMANAGER"
            xfrommail = Trim(rsx1(0))
        End If

    End If
    Call ofv.cerrarconsulta(rsx1)



    StrSql = "Select email,DESCRIPCION From Usuarios  where ID = '" & xusr & "' "

    Set rsx11 = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
    If Not rsx11.EOF Then
        If Len(rsx11(0)) > 0 Then
            xfromusr = "COREMANAGER"
            xfrom = Trim(rsx11(0))
            XPARA2 = "DESTINATARIO " & rsx11(1)
        End If
    End If
    Call ofv.cerrarconsulta(rsx11)
    Call ofv.cerrarconn(xcn)

    '''  XXXMAIL

    If (xfrom <> "" Or Len(xfrom) > 0) Then

        ' message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & xmsgusrn & "</font><b>"
        message = "<P " & xstyle & " >" & XPARA2 & "<BR><BR><b>" & xmsgusrn & "<b></P>"
        Call SendMail(message, xfrom, xfrommail, "CIRCUITO DE AUTORIZACION", session("MailComponent"), "SI", 1, "NO")

    End If

    Set xcn = ofv.conectar(ofv.strconn2)

    xnovcod = xsnovcodi
    dFechai = CLng(Date) - nFecini
    dFechae = CLng(Date) - nFecini + 300
    xhora = Time()

    StrSql = "Insert into novedades (codobjeto,tipobjeto,codigo,tipo,fecinicio,fecexp,mensaje,importancia,relacion,workflowid,contacto,hora) "
    StrSql = StrSql & " values (" & XNCODIGO & ",'" & XSTIPO & "','" & xnovcod & "',4," & dFechai & "," & dFechae & ",'" & xmsgusrn & "',0,0,0,'" & xfromusr & "','" & xhora & "')"
    Call ofv.crearconsultaEx(StrSql, xcn, 1, parametros)

    StrSql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
    StrSql = StrSql & " values ('" & xnovcod & "','" & xusr & "',0," & dFechae & ",'" & xhora & "')"
    Call ofv.crearconsultaEx(StrSql, xcn, 1, parametros)

    Call ofv.cerrarconn(xcn)

End Function



Function GETWKFMANAGER(XNCODIGO, XSTIPO, XSESTADO1, XTABLA, XCKT)


    Set CNXM = ofv.conectar(ofv.strconn2)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''' circuitos de autoevaluacion - riesgo operacional
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = DESARROLLO
    ''''
    If UCase(XTABLA) = "ORRSKCIRCUITO" And XSESTADO1 = 1 Then
        '              response.write "entro aca "

        If XCKT = "WKF" Then
            StrSql = "Update ORRSKCIRCUITO Set "
            StrSql = StrSql & " estado = 1 Where OBJID = 0" & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        End If

        If XCKT = "SMP" Then
            dFechai = CLng(Date) - nFecini
            StrSql = "SELECT codigo FROM novedades "
            StrSql = StrSql & " Where fecexp > " & dFechai
            StrSql = StrSql & " And tipo = 4"
            StrSql = StrSql & " And codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            '       response.write "0 " & strsql & "<br>"
            Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
            If Not rsx.EOF Then
                Do Until rsx.EOF

                    StrSql = "update novedades set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "1 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)



                    StrSql = "update novedadesdist set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "2 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                    rsx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rsx)

        End If



        '           elimina los autorizantes por default cuando pasa a desarrollo nuevamente

        StrSql = "delete from objwfldet where orden > 8999 and "
        StrSql = StrSql & " codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)



    End If

    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = DESARROLLO
    ''''


    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = OBSERVADO
    ''''
    If UCase(XTABLA) = "ORRSKCIRCUITO" And XSESTADO1 = 6 Then
        '              response.write "entro aca "

        If XCKT = "WKF" Then

            StrSql = "Update ORRSKCIRCUITO Set "
            StrSql = StrSql & " estado = 6 Where OBJID = 0" & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

        End If


        If XCKT = "SMP" Then
            dFechai = CLng(Date) - nFecini
            StrSql = "SELECT codigo FROM novedades "
            StrSql = StrSql & " Where fecexp > " & dFechai
            StrSql = StrSql & " And tipo = 4"
            StrSql = StrSql & " And codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            '       response.write "0 " & strsql & "<br>"
            Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
            If Not rsx.EOF Then
                Do Until rsx.EOF

                    StrSql = "update novedades set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "1 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)



                    StrSql = "update novedadesdist set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "2 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                    rsx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rsx)

        End If

        '           elimina los autorizantes por default cuando pasa a desarrollo nuevamente

        StrSql = "delete from objwfldet where orden > 8999 and "
        StrSql = StrSql & " codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)



    End If

    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = OBSERVADO
    ''''

    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = PARA OBSERVACION
    ''''


    If UCase(XTABLA) = "ORRSKCIRCUITO" And XSESTADO1 = 5 Then



        If XCKT = "WKF" Then
            StrSql = "Update ORRSKCIRCUITO Set "
            StrSql = StrSql & "  estado = 5 Where OBJID = 0" & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        End If


        '           incorpora a los autorizantes por default

        xactiva = 0

        If XCKT = "SMP" Then


            StrSql = "select * from Defwfl "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                xactiva = rx(1)

            End If
            Call ofv.cerrarconsulta(rx)

            If xactiva Then
                StrSql = " select * from objwfldet  "
                StrSql = StrSql & "  INNER  join tipObjeto on (tipobjeto.tipobjeto='" & XSTIPO & "' )"
                StrSql = StrSql & "     Where (objwfldet.codobjeto = 0 OR  objwfldet.TIPOBJETO='ALL') "
                If session("PASOOBS") = "OBS" Then
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAUTO=objwfldet.TIPOBJETO"
                Else
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAPRO=objwfldet.TIPOBJETO"
                End If
                StrSql = StrSql & "  order by orden, objid "

                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF



                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & rx(4) & "','SC','NO')"
                        response.write StrSql
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)


                '-------------------------------------------------------
                'INSERT DEFAULT USUARIOS POR TIPO DE ESQUEMA
' incidente nro 1374
                StrSql = "select * from OBJSDKdet Where codobjeto = 0 And tiposeg ='0' AND TIPOBJETO='" & XSTIPO & "'"
                ' strsql = strsql & " and tipobjeto ='ALL'  order by orden, objid"


                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF


                        XUSUC = Mid(rsx(4) & "", 5, Len(rsx(4)) & "")
                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & XUSUC & "','SC','NO')"

                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)

                '-------------------------------------------------------


            End If



            StrSql = "select * from objwfldet Where codobjeto = "
            StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' order by orden, objid"
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                Do Until rx.EOF

                    xnovcodi = "ENA" & rx(4) & CStr(XNCODIGO) & CStr(CLng(Date)) & CStr(Time())

                    StrSql = "Update Objwfldet Set conf = 'SC', enviado='NO',codigo='" & xnovcodi & "' Where objid = " & rx(0)
                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                    rx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rx)

            smsgchat = EsqnomP(XSTIPO, XNCODIGO, ofv) & ">>>"

            StrSql = "select * from objwfl Where codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                sesqauto = rx(3)
                sordenest = rx(4)
                schat = rx(5)
                smsgchat = smsgchat & rx(7)
            End If
            Call ofv.cerrarconsulta(rx)

            If sesqauto And schat Then

                StrSql = "select usr,codigo,objid from objwfldet Where codobjeto = "
                StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "'  order by orden"
                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF

                        Call msgsend(rx(0), smsgchat, ofv, XNCODIGO, XSTIPO, rx(1))
                        StrSql = "Update Objwfldet Set enviado='SI' Where objid = " & rx(2)
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                        If sordenest Then Exit Do
                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)
            End If

        End If

    End If


    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = PARA observacion
    ''''

    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = PARA APROBACION
    ''''


    If UCase(XTABLA) = "ORRSKCIRCUITO" And XSESTADO1 = 2 Then

        If XCKT = "WKF" Then
            StrSql = "Update ORRSKCIRCUITO Set "
            StrSql = StrSql & " estado = 2 Where OBJID = 0" & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        End If


        '           incorpora a los autorizantes por default

        xactiva = 0

        If XCKT = "SMP" Then


            StrSql = "select * from Defwfl "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                xactiva = rx(1)

            End If
            Call ofv.cerrarconsulta(rx)

            If xactiva Then
                StrSql = " select * from objwfldet  "
                StrSql = StrSql & "  INNER  join tipObjeto on (tipobjeto.tipobjeto='" & XSTIPO & "' )"
                StrSql = StrSql & "     Where (objwfldet.codobjeto = 0 OR  objwfldet.TIPOBJETO='ALL') "
                If session("PASOOBS") = "OBS" Then
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAUTO=objwfldet.TIPOBJETO"
                Else
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAPRO=objwfldet.TIPOBJETO"
                End If
                StrSql = StrSql & "  order by orden, objid "

                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF



                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & rx(4) & "','SC','NO')"
                        response.write StrSql
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)


                '-------------------------------------------------------
                'INSERT DEFAULT USUARIOS POR TIPO DE ESQUEMA
' incidente nro 1374
                StrSql = "select * from OBJSDKdet Where codobjeto = 0 And tiposeg ='0' AND TIPOBJETO='" & XSTIPO & "'"
                ' strsql = strsql & " and tipobjeto ='ALL'  order by orden, objid"


                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF


                        XUSUC = Mid(rsx(4) & "", 5, Len(rsx(4)) & "")
                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & XUSUC & "','SC','NO')"

                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)

                '-------------------------------------------------------


            End If



            StrSql = "select * from objwfldet Where codobjeto = "
            StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' order by orden, objid"
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                Do Until rx.EOF

                    xnovcodi = "ENA" & rx(4) & CStr(XNCODIGO) & CStr(CLng(Date)) & CStr(Time())

                    StrSql = "Update Objwfldet Set conf = 'SC', enviado='NO',codigo='" & xnovcodi & "' Where objid = " & rx(0)
                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                    rx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rx)

            smsgchat = EsqnomP(XSTIPO, XNCODIGO, ofv) & ">>>"

            StrSql = "select * from objwfl Where codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                sesqauto = rx(3)
                sordenest = rx(4)
                schat = rx(5)
                smsgchat = smsgchat & rx(7)
            End If
            Call ofv.cerrarconsulta(rx)

            If sesqauto And schat Then

                StrSql = "select usr,codigo,objid from objwfldet Where codobjeto = "
                StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' order by orden"
                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF

                        Call msgsend(rx(0), smsgchat, ofv, XNCODIGO, XSTIPO, rx(1))
                        StrSql = "Update Objwfldet Set enviado='SI' Where objid = " & rx(2)
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                        If sordenest Then Exit Do
                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)
            End If

        End If

    End If


    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = PARA APROBACION
    ''''

    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = VIGENTE
    ''''


    If UCase(XTABLA) = "ORRSKCIRCUITO" And XSESTADO1 = 3 Then

        vFECHAvig = 0
        StrSql = "SELECT * FROM ORRSKCIRCUITO  "
        StrSql = StrSql & " Where OBJID = 0" & XNCODIGO
        Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
        If Not rx.EOF Then
            vFECHAvig = rx("fecvigen")
            If vFECHAvig = "" Then vFECHAvig = 0
        End If
        Call ofv.cerrarconsulta(rx)

        vFECHAHOY = CLng(Date) - CLng(CDate("01-01-1900"))

        StrSql = "Update ORRSKCIRCUITO Set estado = 3 "

        If vFECHAvig = 0 Then
            vFECHAvig = vFECHAHOY
        End If

        StrSql = StrSql & " , fecvigen = '" & vFECHAvig & "' "
        StrSql = StrSql & " Where OBJID = 0" & XNCODIGO
        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)


    End If

    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = VIGENTE
    ''''


    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = NO VIGENTE
    ''''


    If UCase(XTABLA) = "ORRSKCIRCUITO" And XSESTADO1 = 4 Then


        vFECHAHOY = CLng(Date) - CLng(CDate("01-01-1900"))

        StrSql = "Update ORRSKCIRCUITO Set estado = 4 "
        StrSql = StrSql & " , FECBAJA = '" & vFECHAHOY & "' "
        StrSql = StrSql & " Where OBJID = 0" & XNCODIGO
        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)


    End If

    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = NO VIGENTE
    ''''



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''' esquemas normativos - emanager
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = DESARROLLO
    ''''
    If UCase(XTABLA) = "OBJPRTY" And XSESTADO1 = 1 Then
        '              response.write "entro aca "

        If XCKT = "WKF" Then
            StrSql = "Update Objprty Set aprobo = '"
            StrSql = StrSql & "' , estado = 1 Where codobjeto = " & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        End If


        If XCKT = "SMP" Then
            dFechai = CLng(Date) - nFecini
            StrSql = "SELECT codigo FROM novedades "
            StrSql = StrSql & " Where fecexp > " & dFechai
            StrSql = StrSql & " And tipo = 4"
            StrSql = StrSql & " And codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            '       response.write "0 " & strsql & "<br>"
            Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
            If Not rsx.EOF Then
                Do Until rsx.EOF

                    StrSql = "update novedades set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "1 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)



                    StrSql = "update novedadesdist set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "2 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                    rsx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rsx)

        End If

        '           elimina los autorizantes por default cuando pasa a desarrollo nuevamente

        StrSql = "delete from objwfldet where orden > 8999 and "
        StrSql = StrSql & " codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)



    End If

    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = DESARROLLO
    ''''

    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = OBSERVADO
    ''''
    If UCase(XTABLA) = "OBJPRTY" And XSESTADO1 = 6 Then
        '              response.write "entro aca "

        If XCKT = "WKF" Then
            StrSql = "Update Objprty Set aprobo = '"
            StrSql = StrSql & "' , estado = 6 Where codobjeto = " & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        End If


        If XCKT = "SMP" Then
            dFechai = CLng(Date) - nFecini
            StrSql = "SELECT codigo FROM novedades "
            StrSql = StrSql & " Where fecexp > " & dFechai
            StrSql = StrSql & " And tipo = 4"
            StrSql = StrSql & " And codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            '       response.write "0 " & strsql & "<br>"
            Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
            If Not rsx.EOF Then
                Do Until rsx.EOF

                    StrSql = "update novedades set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "1 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)



                    StrSql = "update novedadesdist set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                    '               response.write "2 " & strsql & "<br>"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                    rsx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rsx)

        End If

        '           elimina los autorizantes por default cuando pasa a desarrollo nuevamente

        StrSql = "delete from objwfldet where orden > 8999 and "
        StrSql = StrSql & " codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)



    End If

    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = OBSERVADO
    ''''


    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = PARA OBSERVACION
    ''''


    If UCase(XTABLA) = "OBJPRTY" And XSESTADO1 = 5 Then



        If XCKT = "WKF" Then
            StrSql = "Update Objprty Set aprobo = '"
            StrSql = StrSql & "' , estado = 5 Where codobjeto = " & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        End If


        '           incorpora a los autorizantes por default

        xactiva = 0

        If XCKT = "SMP" Then


            StrSql = "select * from Defwfl "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                xactiva = rx(1)

            End If
            Call ofv.cerrarconsulta(rx)

            If xactiva Then
                StrSql = " select * from objwfldet  "
                StrSql = StrSql & "  INNER  join tipObjeto on (tipobjeto.tipobjeto='" & XSTIPO & "' )"
                StrSql = StrSql & "     Where (objwfldet.codobjeto = 0 OR  objwfldet.TIPOBJETO='ALL') "
                If session("PASOOBS") = "OBS" Then
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAUTO=objwfldet.TIPOBJETO"
                Else
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAPRO=objwfldet.TIPOBJETO"
                End If
                StrSql = StrSql & "  order by orden, objid "

                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF



                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & rx(4) & "','SC','NO')"
                        response.write StrSql
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)


                '-------------------------------------------------------
                'INSERT DEFAULT USUARIOS POR TIPO DE ESQUEMA
' incidente nro 1374
                StrSql = "select * from OBJSDKdet Where codobjeto = 0 And tiposeg ='0'  AND TIPOBJETO='" & XSTIPO & "'"
                ' strsql = strsql & " and tipobjeto ='ALL'  order by orden, objid"


                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF


                        XUSUC = Mid(rsx(4) & "", 5, Len(rsx(4)) & "")
                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & XUSUC & "','SC','NO')"

                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)

                '-------------------------------------------------------


            End If



            StrSql = "select * from objwfldet Where codobjeto = "
            StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' order by orden, objid"
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                Do Until rx.EOF

                    xnovcodi = "ENA" & rx(4) & CStr(XNCODIGO) & CStr(CLng(Date)) & CStr(Time())

                    StrSql = "Update Objwfldet Set conf = 'SC', enviado='NO',codigo='" & xnovcodi & "' Where objid = " & rx(0)
                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                    rx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rx)

            smsgchat = EsqnomP(XSTIPO, XNCODIGO, ofv) & ">>>"

            StrSql = "select * from objwfl Where codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                sesqauto = rx(3)
                sordenest = rx(4)
                schat = rx(5)
                smsgchat = smsgchat & rx(7)
            End If
            Call ofv.cerrarconsulta(rx)

            If sesqauto And schat Then

                StrSql = "select usr,codigo,objid from objwfldet Where codobjeto = "
                StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' order by orden"
                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF

                        Call msgsend(rx(0), smsgchat, ofv, XNCODIGO, XSTIPO, rx(1))
                        StrSql = "Update Objwfldet Set enviado='SI' Where objid = " & rx(2)
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                        If sordenest Then Exit Do
                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)
            End If

        End If

    End If


    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = PARA observacion

    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = PARA APROBACION
    ''''


    If UCase(XTABLA) = "OBJPRTY" And XSESTADO1 = 2 Then

        If XCKT = "WKF" Then
            StrSql = "Update Objprty Set aprobo = '"
            StrSql = StrSql & "' , estado = 2 Where codobjeto = " & XNCODIGO
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        End If


        '           incorpora a los autorizantes por default

        xactiva = 0

        If XCKT = "SMP" Then


            StrSql = "select * from Defwfl "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                xactiva = rx(1)

            End If
            Call ofv.cerrarconsulta(rx)

            If xactiva Then
                '  strsql = "select * from objwfldet Where codobjeto = 0"
                '  strsql = strsql & " and tipobjeto ='ALL'  order by orden, objid"


                StrSql = " select * from objwfldet  "
                StrSql = StrSql & "  INNER  join tipObjeto on (tipobjeto.tipobjeto='" & XSTIPO & "' )"
                StrSql = StrSql & "     Where (objwfldet.codobjeto = 0 OR  objwfldet.TIPOBJETO='ALL') "
                If session("PASOOBS") = "OBS" Then
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAUTO=objwfldet.TIPOBJETO"
                Else
                    StrSql = StrSql & " AND NOT  tipObjeto.WKFAPRO=objwfldet.TIPOBJETO"
                End If
                StrSql = StrSql & "  order by orden, objid "

                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF



                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & rx(4) & "','SC','NO')"
                        response.write StrSql
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)


                '-------------------------------------------------------
                'INSERT DEFAULT USUARIOS POR TIPO DE ESQUEMA
' incidente nro 1374
                StrSql = "select * from OBJSDKdet Where codobjeto = 0 And tiposeg ='0' AND TIPOBJETO='" & XSTIPO & "'"
                ' strsql = strsql & " and tipobjeto ='ALL'  order by orden, objid"


                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF


                        XUSUC = Mid(rsx(4) & "", 5, Len(rsx(4)) & "")
                        StrSql = "insert into Objwfldet (codobjeto, tipobjeto, orden,usr,conf, enviado) "
                        StrSql = StrSql & " values(" & XNCODIGO & ",'" & XSTIPO & "'," & CLng(rx(3)) + 9000 & ",'"
                        StrSql = StrSql & XUSUC & "','SC','NO')"

                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)

                '-------------------------------------------------------


            End If



            StrSql = "select * from objwfldet Where codobjeto = "
            StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' order by orden, objid"
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                Do Until rx.EOF

                    xnovcodi = "ENA" & rx(4) & CStr(XNCODIGO) & CStr(CLng(Date)) & CStr(Time())

                    StrSql = "Update Objwfldet Set conf = 'SC', enviado='NO',codigo='" & xnovcodi & "' Where objid = " & rx(0)
                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)

                    rx.movenext
                Loop
            End If
            Call ofv.cerrarconsulta(rx)

            smsgchat = EsqnomP(XSTIPO, XNCODIGO, ofv) & ">>>"

            StrSql = "select * from objwfl Where codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
            Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
            If Not rx.EOF Then
                sesqauto = rx(3)
                sordenest = rx(4)
                schat = rx(5)
                smsgchat = smsgchat & rx(7)
            End If
            Call ofv.cerrarconsulta(rx)

            If sesqauto And schat Then

                StrSql = "select usr,codigo,objid from objwfldet Where codobjeto = "
                StrSql = StrSql & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' order by orden"
                Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                If Not rx.EOF Then
                    Do Until rx.EOF

                        Call msgsend(rx(0), smsgchat, ofv, XNCODIGO, XSTIPO, rx(1))
                        StrSql = "Update Objwfldet Set enviado='SI' Where objid = " & rx(2)
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                        If sordenest Then Exit Do
                        rx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rx)
            End If

        End If

    End If


    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = PARA APROBACION
    ''''

    ''''
    '''' WKFLOW --------------- POR CAMBIO ESTADO = VIGENTE
    ''''


    If UCase(XTABLA) = "OBJPRTY" And XSESTADO1 = 3 Then

        vFECHAvig = 0
        StrSql = "SELECT * FROM OBJprty  "
        StrSql = StrSql & " Where codobjeto=" & XNCODIGO
        StrSql = StrSql & " And tipobjeto = '" & XSTIPO & "'"
        Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
        If Not rx.EOF Then
            vFECHAvig = rx("fecvigen")
            If vFECHAvig = "" Then vFECHAvig = 0
        End If
        Call ofv.cerrarconsulta(rx)

        vFECHAHOY = CLng(Date) - CLng(CDate("01-01-1900"))

        StrSql = "Update Objprty Set aprobo = '" & session("usuario")
        StrSql = StrSql & "' , estado = 3 "

        If vFECHAvig = 0 Then
            vFECHAvig = vFECHAHOY
        End If

        StrSql = StrSql & " , fecvigen = '" & vFECHAvig & "' , fecpublic = '" & vFECHAHOY & "' "
        StrSql = StrSql & " Where codobjeto = " & XNCODIGO
        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)





        Sql = "Select objrem,Tipobjeto From Objetos Where codobjeto = " & XNCODIGO
        Set rsCodigo = ofv.crearconsultaEx(Sql, CNXM, 1, volver)
        If Not rsCodigo.EOF Then
            If CLng(rsCodigo(0)) <> 0 Then
                Sql = "Select Tipobjeto From Objetos Where codobjeto = " & rsCodigo(0)
                Set rsCodigox = ofv.crearconsultaEx(Sql, CNXM, 1, volver)
                If Not rsCodigox.EOF Then xntipobj = rsCodigox(0)
                '-----------------
                Segct = "S"

                Sql = "Select * From OBJSDK Where tipobjeto='"
                Sql = Sql & xntipobj & "' and codobjeto = " & rsCodigo(0)
                Set rsx = ofv.crearconsultaEx(Sql, CNXM, 1, parametros)
                If Not rsx.EOF Then


                    If session("Usuario") <> rsx(7) Then
                        If rsx(3) Then Segct = "N"
                        Call ofv.cerrarconsulta(rsx)
                        Sql = "Select * From OBJSDKDET Where tipobjeto='" & xntipobj
                        Sql = Sql & "' and codobjeto = " & rsCodigo(0)
                        Sql = Sql & " and ((grpusr = 'USR-" & session("Usuario") & "') "
                        Sql = Sql & " or (grpusr = 'GRP-" & session("Seggrp") & "')) "
                        Set rsx = ofv.crearconsultaEx(Sql, CNXM, 1, parametros)
                        Do Until rsx.EOF

                            Select Case rsx(3)
                                Case "1"
                                    Segct = "S"
                            End Select
                            rsx.movenext
                        Loop
                    End If
                End If
                Call ofv.cerrarconsulta(rsx)

                '-----------
                If Segct = "S" Then
                    vFECHAHOY = CLng(Date) - CLng(CDate("01-01-1900"))

                   dim Areempaza()
                   ReDim  Areempaza(3,0)
                   Areempaza(0,0)= rsCodigo(0)
                   Areempaza(1,0)= "S"
                   Areempaza(2,0)= "S"
				   '  incidente 1353
					Areempaza(3,0)=rsCodigo(1)
									 '  incidente 1353
					ssql="select depenreemp.*,objetos.tipobjeto from depenreemp inner join objetos on (objetos.codobjeto=depenreemp.codobjagr) where codobj=" & XNCODIGO
                 '  response.write ssql
                   Set rsxremp = ofv.crearconsultaEx(sSql, CNXM, 1, parametros)
                   YCRem = 1
                   do while not rsxremp.eof
                    ReDim Preserve Areempaza(2,YCRem)
                    Areempaza(0,YCRem)= rsxremp(3)
				    Areempaza(1,YCRem)= rsxremp(4)
                    Areempaza(2,YCRem)= rsxremp(5)
                    Areempaza(3,YCRem)= rsxremp("tipobjeto")

                    YCRem=YCRem+1
                    rsxremp.movenext
                   loop
                  ' response.write ubound(Areempaza,2)
                     tipobjetop=rsCodigo(1)
                    Call ofv.cerrarconsulta(rsxremp)
                   for IReemplaza=0 to ubound(Areempaza,2)
					'  incidente 1353
                     call Trata_reemplaza(IReemplaza,CNXM,vFECHAHOY,tipobjetop,XNCODIGO,xntipobj,Areempaza(0,IReemplaza),Areempaza(1,IReemplaza),Areempaza(2,IReemplaza),volver,Areempaza(3,IReemplaza))
                   next
                end if


            End If

        End If
        Call ofv.cerrarconsulta(rsCodigo)

        ' envio de mensajes segun lista de distribucion



        Set cnx = ofv.conectar(ofv.strconn0)


        xlista = ""
        xsepar = ""
        xfromusr = "NO"
        xfromusrD = "NO"
        xfrom = "CorESolutionSuite@hotmail.com"
        xAviso = "NO"
        ximport = 1
        xmensaje1 = EsqnomP(XSTIPO, XNCODIGO, ofv) & ">>>"
        xsubject = "(Cor E-Manager) Activacion de " & xmensaje1
        xmensaje = xmensaje1 & " esta activo a partir de la fecha "


        XUSRPRV = session("USRPRV")
        session("USRPRV") = "I"

        XXXX = getInfoUserw("COREMANAGER")
        If session("SGU_mail") <> "" Then
            xfrom = session("SGU_mail")
            xfromusr = session("SGU_nombre")
            xfromusrD = "COREMANAGER"
        End If

        session("USRPRV") = XUSRPRV



        xAviso = "NO"
        ximport = 1
        xdiasvig = 7

        StrSql = "SELECT * FROM OBJdistPrty "
        StrSql = StrSql & " Where codobjeto=" & XNCODIGO
        StrSql = StrSql & " And tipobjeto = '" & XSTIPO & "'"
        Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
        If Not rsx.EOF Then
            xAviso = rsx(7)
            ximport = rsx(6)
            xmensaje = rsx(5)

            '''''''''''''' agregado para Banco Finansur
            If session("APLMDCLI") = "FINANSUR" Then
                xmensajen = Replace(xmensaje, "<ESQNOR>", "<b>" & xmensaje1 & "</b>")
            Else
                xmensajen = Replace(xmensaje, "<ESQNOR>", xmensaje1)
            End If
            '''''''''''''' fin agregado para Banco Finansur

            '''''''''''''' agregado para Banco Finansur
            If session("APLMDCLI") = "FINANSUR" Then
                If session("CORUNQ") = "S" Then
                    xmensurl = getobjurl(XNCODIGO, XSTIPO)
                    xmensaje = Replace(xmensaje, "<ESQNOR>", " <a href='" & xmensurl & "' >" & "<b>" & xmensaje1 & "</b>" & "</a>")
                Else
                    xmensaje = Replace(xmensaje, "<ESQNOR>", "<b>" & xmensaje1 & "</b>")
                End If
            Else
                If session("CORUNQ") = "S" Then
                    xmensurl = getobjurl(XNCODIGO, XSTIPO)
                    xmensaje = Replace(xmensaje, "<ESQNOR>", " <a href='" & xmensurl & "' >" & xmensaje1 & "</a>")
                Else
                    xmensaje = Replace(xmensaje, "<ESQNOR>", xmensaje1)
                End If
            End If
            '''''''''''''' fin agregado para Banco Finansur

            xmensaje = Replace(xmensaje, ">>>", " ")

            xmensajen = Replace(xmensajen, ">>>", "<br>")
            xsubject = rsx(4)
            xsubject = Replace(xsubject, "<ESQNOR>", xmensaje1)
            xsubject = Replace(xsubject, ">>>", " ")
            ' xfromusr  = rsx(3)
            xfromusrD = rsx(3)
            xdiasvig = rsx(8)
        End If
        Call ofv.cerrarconsulta(rsx)

        xnovcod = "ENV" & CStr(XNCODIGO)
        dFechai = CLng(Date) - CLng(CDate("01-01-1900"))
        dFechae = CLng(Date) - CLng(CDate("01-01-1900")) + CLng(xdiasvig)

        Select Case CLng(ximport)
            Case 0
                ximportn = 2
            Case 1
                ximportn = 1
            Case 2
                ximportn = 0

        End Select


        xhora = Time()




        ''''''''''''''''''''''''' BUSCA MAIL Y NOMBRE DEL EMISOR EN CASO DE NO SER EL DEFAULT '''''''''''''''''''''''''


        If session("WKFETPU") = "D" Then               ' EMISOR DEFAULT (EL RESPONSABLE DEL PASO)



            If xfromusrD <> "NO" Then
                If xfromusrD <> "COREMANAGER" Then

                    XXXX = getInfoUserw(xfromusrD)
                    If session("SGU_mail") <> "" Then
                        xfrom = session("SGU_mail")
                        xfromusr = session("SGU_nombre")
                        ''' xfromusrD = "COREMANAGER"
                    End If

                End If
            End If



        ElseIf session("WKFETPU") = "U" Then           ' EMISOR PREDETERMINADO (PARAMETRIZADO EL LA BASE DE USUARIOS)

            WKFEUSR = session("WKFEUSR")

            If WKFEUSR = "" Then WKFEUSR = "COREMANAGER"

            If WKFEUSR <> "COREMANAGER" Then


                XUSRPRV = session("USRPRV")
                session("USRPRV") = "I"

                XXXX = getInfoUserw(WKFEUSR)
                If session("SGU_mail") <> "" Then
                    xfrom = session("SGU_mail")
                    xfromusr = session("SGU_nombre")
                    xfromusrD = session(WKFEUSR)
                End If

                session("USRPRV") = XUSRPRV


            Else
                xfromusrD = "COREMANAGER"

            End If


        Else                                           ' EMISOR FIJO (USUARIO Y MAIL FIJOS EN CORCONECT)

            WKFEUSR = session("WKFEUSR")

            If WKFEUSR = "" Then WKFEUSR = "COREMANAGER"

            xfrom = session("WKFEMAI")
            xfromusr = WKFEUSR
            xfromusrD = WKFEUSR
        End If


        ''''''''''''''''''''''''' ENVIA EL MENSAJE A LA BANDEJA DE NOVEDADES DE ACTIVACION/PUBLICACION DE DOCUMENTOS '''''''''


        StrSql = "Insert into novedades (codobjeto,tipobjeto,codigo,tipo,fecinicio,fecexp,mensaje,importancia,relacion,workflowid,contacto,hora) "
        StrSql = StrSql & " values (" & XNCODIGO & ",'" & XSTIPO & "','" & xnovcod & "',3," & dFechai & "," & dFechae & ",'" & xmensajen & "'," & ximportn & ",0,0,'" & xfromusrD & "','" & xhora & "')"


        Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        StrSql = "SELECT * FROM OBJdist "
        StrSql = StrSql & " Where codobjeto=" & XNCODIGO
        StrSql = StrSql & " And tipobjeto = '" & XSTIPO & "'"
        Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
        If rsx.EOF Then

            StrSql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
            StrSql = StrSql & " values ('" & xnovcod & "','TODOS',999," & dFechae & ",'" & xhora & "')"
            '    response.write strsql & "<br>"
            Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
        Else

            If session("APLMDCLI") = "BPN" Then
                StrSql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
                StrSql = StrSql & " values ('" & xnovcod & "','TODOS',999," & dFechae & ",'" & xhora & "')"
                '    response.write strsql & "<br>"
                Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
            End If
            Do Until rsx.EOF

                If UCase(rsx(4)) = "SECT-" Then

                    StrSql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,xhora) "
                    StrSql = StrSql & " values ('" & xnovcod & "','SECTOR',0" & rsx(5) & "," & dFechae & ",'" & xhora & "')"

                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)

                    StrSql = "Select email From Usuarios  where sector = 0" & rsx(5)
                    Set rsx1 = ofv.crearconsultaEx(StrSql, cnx, 1, volver)
                    Do Until rsx1.EOF

                        xmail = rsx1("email")
                        If Len(xmail) > 0 Then
                            xmail = Trim(xmail)
                            xlista = xlista & xsepar & xmail
                            xsepar = " ; "
                        End If


                        rsx1.movenext
                    Loop
                    Call ofv.cerrarconsulta(rsx1)






                Else

                    If rsx(4) = "TODOS" Then
                        XSECVAL = rsx(5)
                    Else
                        XSECVAL = 0
                    End If
                    If Mid(rsx(4) & "", 1, 4) = "GRM-" Then

                        sgrupoM = Mid(rsx(4) & "", 5, Len(rsx(4)) & "")
                        ' session("SGU_mail")=""
                        xxx = getMailDelGrupo(sgrupoM)
                        ' respons.write xxx
                        session("SGU_mail") = xxx

                        If session("SGU_mail") <> "" Then
                            xlista = xlista & xsepar & session("SGU_mail")
                            xsepar = " ; "
                        End If
                        If session("APLMDCLI") = "XBPN" Then
                            StrSql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
                            StrSql = StrSql & " values ('" & xnovcod & "','TODOS',999," & dFechae & ",'" & xhora & "')"
                            '    response.write strsql & "<br>"
                            Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                        End If
                    Else
                        StrSql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
                        StrSql = StrSql & " values ('" & xnovcod & "','" & rsx(4) & "',0" & XSECVAL & "," & dFechae & ",'" & xhora & "')"

                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                        XXXX = getInfoUserw(rsx(4))
                        If session("SGU_mail") <> "" Then
                            xlista = xlista & xsepar & session("SGU_mail")
                            xsepar = " ; "
                        End If
                    End If

                End If



                rsx.movenext
            Loop
        End If
        Call ofv.cerrarconsulta(rsx)
        Call ofv.cerrarconn(cnx)

        ''''''''''''''''''''''' envia el mail a la lista XXXMAIL

        If Len(xlista) > 0 Then


            '''' "Circuito de autorización"

            If session("APLMDCLI") = "LACAJA" Or session("APLMDCLI") = "CFA" Then
                XPARA2 = ""
                message = "<DIV STYLE='background-color:" & session("CMFMENU") & ";height:40;width:600px;"
                message = message & "BORDER-color:#E0E0E0;BORDER-STYLE:SOLID;BORDER-WIDTH:1;'  >"
                ' Whtml = Whtml & "<br>"
                message = message & "<P VALIGN=MIDDLE >&nbsp;&nbsp;<font size=4 face='" & session("CMFMEFTFC") & "' color='" & session("CMFMEFT") & "' ><B>"
                If session("APLMDCLI") = "LACAJA" Then
                    message = message & "NORMAS Y REINGENIERIA DE PROCESOS" & "</B></FONT></P>"
                Else
                    message = message & session("CorTitApli") & "</B></FONT></P>"
                End If
                message = message & "</DIV>"
                message = message & "<BR>"

                message = message & "<DIV STYLE='BORDER-color:" & session("CMFMENU") & ";BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                message = message & "height:50;width:600px;'  >"
                message = message & "<CENTER>"



                message = message & "<br>"
                message = message & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                message = message & "background-color:#F0F0F0;height:25;width:100%;'  >"
                message = message & "<P ALIGN=LEFT >"
                message = message & "<font size=1 color=#888888 face='" & session("CMFMEFTFC") & "' ><B>"
                message = message & "&nbsp;&nbsp;DESCRIPCION</B></FONT></P>"
                message = message & "</DIV>"

                message = message & "<br>"
                message = message & "<DIV STYLE='width:80%;' >"
                message = message & "<P ALIGN='JUSTIFY' >"
                message = message & "<font size=2 color=#888888 face='" & session("CMFMEFTFC") & "' >"
                message = message & Replace(xmensaje & "", Chr(13), "<br>")
                XMESAURL = ""
                If session("APLMDCLI") = "LACAJA" Then
                    If session("URLTYP") = "V" Then

                        XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))

                    End If

                    If session("URLPORT") = "N" Then
                        XMESAURL = "http://" & request.servervariables("SERVER_NAME") & XMESAURL
                    Else
                        XMESAURL = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
                    End If
                    XMESAURL = "<br><br><a href='" & XMESAURL & "'>INGRESAR A LA APLICACIÓN</a>"
                End If
                message = message & XMESAURL

                message = message & "</FONT></P>"
                message = message & "</DIV>"
                message = message & "<BR>"
                message = message & "</DIV>"

                message = message & "<br>"
                message = message & "<DIV ALIGN='RIGHT' STYLE='width:200px;' >"
                message = message & "<P  >"
                message = message & "<font size=1 color=#6699cc face='" & session("CMFMEFTFC") & "' ><B><i>"
                message = message & "COR&nbsp;E-SOLUTION&nbsp;SUITE&nbsp;V4.8.5</i></B></FONT></P>"
                message = message & "</DIV>"

            Else
                XPARA2 = "DESTINATARIO LISTA DE DISTRIBUCION"

                '' modificacion exclusiva para Banco Finansur

                If session("APLMDCLI") = "FINANSUR" Then
                    message = "<P " & xstyle & ">" & XPARA2 & "<BR><BR>" & xmensaje & "</P>"
                Else
                    message = "<P " & xstyle & ">" & XPARA2 & "<BR><BR><b>" & xmensaje & "<b></P>"
                End If
            End If
            'message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & xmensaje & "</font><b>"





            Call SendMail(message, xlista, xfrom, xsubject, session("MailComponent"), "SI", ximport, xAviso)

        End If


    End If





    ''  END IF

    ''''
    '''' WKFLOW --------------- FIN ---- POR CAMBIO ESTADO = VIGENTE
    ''''


    ''''
    '''' WKFLOW --------------- POR CAMBIO CIRCUITO SIMPLE DE APROBACION
    ''''


    If UCase(XTABLA) = "OBJWFLDET" Then

        smsgchat = EsqnomP(XSTIPO, XNCODIGO, ofv) & ">>>"
        StrSql = "select * from objwfl Where codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
        Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
        If Not rx.EOF Then
            sesqauto = rx(3)
            sordenest = rx(4)
            schat = rx(5)
            smsgchat = smsgchat & " " & rx(7)
        End If
        Call ofv.cerrarconsulta(rx)
        sant = ""
        sdat = ""
        sdatSC = ""
        sdatNC = ""
        sdatNCID = 0
        susr = ""
        StrSql = "SELECT * FROM OBJwfldet  "
        StrSql = StrSql & " Where codobjeto=" & XNCODIGO
        StrSql = StrSql & " And tipobjeto = '" & XSTIPO & "' order by orden,objid"

        Set rx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
        Do Until rx.EOF

            If rx(8) = "SC" Then
                susr = rx(4)
                sdat = rx(8)
                sdatSC = rx(8)
                sdatNCID = rx(0)
                sdatNC = rx(10)
                If sesqauto And sordenest Then Exit Do
            ElseIf rx(8) = "NO" Then
                susr = rx(4)
                sdat = rx(8)
                Exit Do
            ElseIf rx(8) = "NC" Then
                susr = rx(4)
                sdat = rx(8)
                sdatNC = rx(8)
                sdatNCID = rx(0)
                If sesqauto And sordenest Then
                    dFechai = CLng(Date) - CLng(CDate("01-01-1900"))
                    StrSql = "SELECT codigo FROM novedades "
                    StrSql = StrSql & " Where fecexp > " & dFechai
                    StrSql = StrSql & " And tipo = 4"
                    StrSql = StrSql & " And codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
                    StrSql = StrSql & " And codigo = '" & rx(10) & "' "

                    Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                    If Not rsx.EOF Then


                        StrSql = "update novedades set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "


                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)



                        StrSql = "update novedadesdist set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "

                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)

                    End If
                    Call ofv.cerrarconsulta(rsx)

                    smsgchatx = "Su asignacion para autorizacion ha sido expirada "
                    Call msgsend(susr, smsgchatx, ofv, XNCODIGO, XSTIPO, rx(10) & "ANL")
                End If


                StrSql = "update OBJwfldet set conf = 'NP', enviado='SP', fecconf = '" & Now() & "',comen1 = '"
                StrSql = StrSql & "Su asignacion para autorizacion ha sido expirada ' where objid = " & rx(0)
                Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)

            ElseIf rx(8) = "NP" Then
                sant = rx(8)
                sdat = rx(8)
            Else
                sant = rx(8)
                sdat = rx(8)
                If sesqauto And sordenest And schat Then
                    If rx(11) = "SI" Then
                        StrSql = "Update Objwfldet Set enviado='SP' Where objid = " & rx(0)
                        Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                        dFechai = CLng(Date) - CLng(CDate("01-01-1900"))
                        StrSql = "SELECT codigo FROM novedades "
                        StrSql = StrSql & " Where fecexp > " & dFechai
                        StrSql = StrSql & " And tipo = 4"
                        StrSql = StrSql & " And codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
                        StrSql = StrSql & " And codigo = '" & rx(10) & "' "

                        Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                        If Not rsx.EOF Then


                            StrSql = "update novedades set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "


                            Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)



                            StrSql = "update novedadesdist set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "

                            Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)

                        End If
                        Call ofv.cerrarconsulta(rsx)
                    End If
                End If
            End If
            rx.movenext
        Loop
        Call ofv.cerrarconsulta(rx)

        Select Case sdat
            Case "SC"
                If sesqauto And sordenest And schat Then
                    Call msgsend(susr, smsgchat, ofv, XNCODIGO, XSTIPO, sdatNC)
                    StrSql = "Update Objwfldet Set enviado='SI' Where objid = " & sdatNCID
                    Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)
                End If

            Case "NO"


                StrSql = "Update Objprty Set aprobo = '"
                StrSql = StrSql & "' , estado = 1 Where codobjeto = " & XNCODIGO
                Call ofv.crearconsultaEx(StrSql, CNXM, 1, volver)




                XXXX = GETWKFMANAGER(XNCODIGO, XSTIPO, 1, "OBJPRTY", XCKT)


            Case Else
                ' entra por todos ok o por si alguno fue anulado por expiracion --- --- --- ---

                If sdatSC <> "SC" Then


                    dFechai = CLng(Date) - CLng(CDate("01-01-1900"))
                    StrSql = "SELECT codigo FROM novedades "
                    StrSql = StrSql & " Where fecexp > " & dFechai
                    StrSql = StrSql & " And tipo = 4"
                    StrSql = StrSql & " And codobjeto = " & XNCODIGO & " AND TIPOBJETO = '" & XSTIPO & "' "
                    '               response.write "0 " & strsql & "<br>"
                    Set rsx = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                    If Not rsx.EOF Then
                        Do Until rsx.EOF

                            StrSql = "update novedades set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                            '              response.write "1 " & strsql & "<br>"

                            Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)



                            StrSql = "update novedadesdist set fecexp = " & dFechai & " where codigo = '" & rsx(0) & "' "
                            '               response.write "2 " & strsql & "<br>"

                            Call ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                            rsx.movenext
                        Loop
                    End If
                    Call ofv.cerrarconsulta(rsx)

                    XXXX = GETWKFMANAGER(XNCODIGO, XSTIPO, 3, "OBJPRTY", XCKT)

                End If
        End Select
    End If

    ''''
    '''' WKFLOW --------------- FIN -------- POR CAMBIO CIRCUITO SIMPLE DE APROBACION
    ''''

    ''  end if

    Call ofv.cerrarconn(CNXM)

End Function


Function getobjurl(getnCodigo, getstipo)


    WKURLCONS = getnCodigo & "-" & getstipo

    Set getcn = ofv.conectar(ofv.strconn2)


    xStrSql = "SELECT * FROM TIPOBJETO  "
    xStrSql = xStrSql & " Where  tipobjeto = '" & getstipo & "' "

    Set getrs = ofv.crearconsultaEx(xStrSql, getcn, 1, parametros)
    If Not getrs.EOF Then

        wkhosturl = request.servervariables("HTTP_REFERER")

        wkhosturlX = UCase(wkhosturl & "NO")

        wkhosturln = InStr(wkhosturlX, "/COREMANAGER/")

        If wkhosturln > 0 Then

            wkhosturl = Left(wkhosturl, wkhosturln)


        Else

            wkhosturln = InStr(UCase(request.servervariables("SCRIPT_NAME") & "NO"), "/COREMANAGER/")

            wkhosturl = Left(request.servervariables("SCRIPT_NAME") & "NO", wkhosturln)

            wkhosturl = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & wkhosturl


        End If

        WKURLCONS = wkhosturl


        WKURLCONS = WKURLCONS & "coress20/corexpress.asp?opt=CC-" & getnCodigo & "-" & getstipo

    End If

    getobjurl = WKURLCONS


End Function


Function getobjurlWK(getnCodigo, getstipo, getpaso)


    WKURLCONS = getnCodigo & "-" & getstipo

    Set getcn = ofv.conectar(ofv.strconn2)


    xStrSql = "SELECT * FROM TIPOBJETO  "
    xStrSql = xStrSql & " Where  tipobjeto = '" & getstipo & "' "

    Set getrs = ofv.crearconsultaEx(xStrSql, getcn, 1, parametros)
    If Not getrs.EOF Then

        wkhosturl = request.servervariables("HTTP_REFERER")

        wkhosturlX = UCase(wkhosturl & "NO")

        wkhosturln = InStr(wkhosturlX, "/COREMANAGER/")

        If wkhosturln > 0 Then

            wkhosturl = Left(wkhosturl, wkhosturln)


        Else

            wkhosturln = InStr(UCase(request.servervariables("SCRIPT_NAME") & "NO"), "/COREMANAGER/")

            wkhosturl = Left(request.servervariables("SCRIPT_NAME") & "NO", wkhosturln)

            wkhosturl = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & wkhosturl


        End If

        WKURLCONS = wkhosturl


        WKURLCONS = WKURLCONS & "coress20/corexpress.asp?opt=WK-" & getnCodigo & "-" & getstipo & "&SEGCKTPSC=" & getpaso

    End If

    getobjurlWK = WKURLCONS


End Function


 '  incidente 1353
sub Trata_reemplaza (IReemplaza,CNXM,vFECHAHOY,rsCodigoT,XNCODIGO,xntipobj,rsCodigoO,slink,sdepen,volver,rsTipoO)

                Sql = "Update Objprty Set estado = 4,fecbaja = '" & vFECHAHOY & "', "
				Sql = Sql & "elimino = '" & sUsuario & "' "
				Sql = Sql & "Where estado = 3 And codobjeto = " & rsCodigoO
				Call ofv.crearconsultaEx(Sql, CNXM, 1, volver)
'response.write sql & "<br>"
                  Sql = "Update Objetos Set objrdo = " & XNCODIGO
                  Sql = Sql & " Where codobjeto =0" &  rsCodigoO
                   Call ofv.crearconsultaEx(Sql, CNXM, 1, volver)
'response.write CNXM.errors.Count & "<br>"
                  If CNXM.errors.Count = 0 Then
                        ' actualiza esquemas asociados (arbol)


                    if slink="S" then
                        Sql = "Update Relaobj Set codobjmaster = " & XNCODIGO & " , tipobjmaster = '"
                        Sql = Sql & rsCodigoT & "' Where codobjmaster = " & rsCodigoO
                        Sql = Sql & " and tipobjmaster = '" & rsTipoO & "' "
                        Call ofv.crearconsultaEx(Sql, CNXM, 1, volver)


                '   response.write Sql & "<br>"



                        Sql = "Update objmsg Set MSGOBJCOD = " & XNCODIGO & " , MSGOBJTIP = '"
                        Sql = Sql & rsCodigoT & "' Where  MSGOBJCOD = " & rsCodigoO
                        'sql = sql & " and tipobjmaster = '" & xntipobj & "' "
                        Call ofv.crearconsultaEx(Sql, CNXM, 1, volver)

                        ''       ' actualiza esquema linkeados
                        ''               sql = "Update objlnk Set codobjetoL = " & XNCODIGO & " , tipobjetoL = '"
                        ''               sql = sql & rsCodigo(1) & "' Where codobjetoL = " & rsCodigo(0)
                        ''               sql = sql & " and tipobjetoL = '" & xntipobj & "' "
                        ''               call ofv.crearconsultaEx(Sql,CNXM,1,volver)
       'response.write Sql
                      end if
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' actualiza esquema linkeados
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



                    if sdepen="S" then


                        Sql = "Select * From objlnk Where tipobjetoL ='"
                        Sql = Sql & xntipobj & "' and codobjetoL = " & rsCodigoO
                        Sql = Sql & " and LTIPO = 'L' and codobjetoL <> codobjeto"
                        ' si es link al mismo doc lo ignora

                        Set RSX2 = ofv.crearconsultaEx(Sql, CNXM, 1, parametros)

                        '''     response.write "-lee link ---" & Strsql & "<br>"

                        Do While Not RSX2.EOF

                    '    select case   IReemplaza

                         ' reemplaza original
                     '   case 0

                            WIDEN3L = 0
                            WIDEN4 = 0
                            WIDEN4L = 0


                            '''''''''''''''''''''' fija la solapa '''''''''''''''''
                            WIDEN3L = RSX2(11)

                            '''''''''''''''''''''' fija el 1er nivel '''''''''''''''''

                            Select Case WIDEN3L

                                Case "2"   'CONSIDERACIONES

                                    StrSql = "SELECT objid FROM OBJNWEDT WHERE relant = '" & RSX2(12) & "' and codobjeto = " & XNCODIGO
                                    Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                    If Not Xrs.EOF Then
                                        WIDEN4 = Xrs(0)
                                    End If
                                    Call ofv.cerrarconsulta(Xrs)


                                Case "5"   'PROCEDIMIENTOS

                                    StrSql = "SELECT CODPROCED FROM objproced WHERE relant = '" & RSX2(12) & "' and codobjeto = " & XNCODIGO
                                    Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                    If Not Xrs.EOF Then
                                        WIDEN4 = Xrs(0)
                                    End If
                                    Call ofv.cerrarconsulta(Xrs)

                                Case "7"   'CONTROLES


                                    StrSql = "SELECT objid FROM objcontrol WHERE relant = '" & RSX2(12) & "' and codobjeto = " & XNCODIGO
                                    Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                    If Not Xrs.EOF Then
                                        WIDEN4 = Xrs(0)
                                    End If
                                    Call ofv.cerrarconsulta(Xrs)


                                Case "8"   'ESQUEMAS

                                    StrSql = "SELECT objid FROM objesquema WHERE relant = '" & RSX2(12) & "' and codobjeto = " & XNCODIGO
                                    Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                    If Not Xrs.EOF Then
                                        WIDEN4 = Xrs(0)
                                    End If
                                    Call ofv.cerrarconsulta(Xrs)


                                Case "9"   'TABLAS


                                    StrSql = "SELECT objid FROM ObjTabDef WHERE relant = '" & RSX2(12) & "' and codobjeto = " & XNCODIGO
                                    Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                    If Not Xrs.EOF Then
                                        WIDEN4 = Xrs(0)
                                    End If
                                    Call ofv.cerrarconsulta(Xrs)


                                Case "11"   'CONTENIDO

                                    StrSql = "SELECT objid FROM OBJESTR WHERE relant = '" & RSX2(12) & "' and codobjeto = " & XNCODIGO
                                    Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                    If Not Xrs.EOF Then
                                        WIDEN4 = Xrs(0)
                                    End If
                                    Call ofv.cerrarconsulta(Xrs)



                            End Select

                            '''      response.write "---- lee anterior ----" & Strsql & "<br>"

                            '''''''''''''''''''''' fija el 2er nivel '''''''''''''''''

                            If WIDEN3L = "9" Then

                                StrSql = "SELECT objid FROM ObjTabblk WHERE relant = '" & RSX2(13) & "' and codobjeto = " & XNCODIGO
                                Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                If Not Xrs.EOF Then
                                    WIDEN4L = Xrs(0)
                                End If
                                Call ofv.cerrarconsulta(Xrs)

                            End If


                            If WIDEN3L = "11" Then

                                StrSql = "SELECT objid FROM OBJESTDET WHERE relant = '" & RSX2(13) & "' and codobjeto = " & XNCODIGO
                                Set Xrs = ofv.crearconsultaEx(StrSql, CNXM, 1, parametros)
                                If Not Xrs.EOF Then
                                    WIDEN4L = Xrs(0)
                                End If
                                Call ofv.cerrarconsulta(Xrs)

                            End If

                            ''''     response.write "---- lee anterior - 2do ---" & Strsql & "<br>"


                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




                            Sql = "Update objlnk Set codobjetoL = " & XNCODIGO
                            Sql = Sql & ", tipobjetoL = '" & rsCodigoT & "' "
                            Sql = Sql & ", IDEN3L     = '" & WIDEN3L & "' "
                            Sql = Sql & ", IDEN4      = '" & WIDEN4 & "' "
                            Sql = Sql & ", IDEN4L     = '" & WIDEN4L & "' "
                            Sql = Sql & " Where OBJID = " & RSX2(0)
                            Call ofv.crearconsultaEx(Sql, CNXM, 1, volver)

                  '  case else
                   '       Sql = "Update objlnk Set codobjetoL = " & XNCODIGO
                  ''        Sql = Sql & ", tipobjetoL = '" & rsCodigoT & "' "
                  '         Sql = Sql & " Where codobjeto = " & RSX2(0)
                   '        Call ofv.crearconsultaEx(Sql, CNXM, 1, volver)

                  '  end select

                            RSX2.movenext

                        Loop


                        Sql = "Update objlnk Set codobjetoL = " & XNCODIGO & " , tipobjetoL = '"
                        Sql = Sql & rsCodigoT & "' Where codobjetoL = " & rsCodigoO
                        Sql = Sql & " and tipobjetoL = '" & rsTipoO & "' "
                       ' Sql = Sql & " and LTIPO <> 'L' and LTIPO <> 'LRGVN' "
                     '   response.write sql & "<br>"
                        Call ofv.crearconsultaEx(Sql, CNXM, 1, volver)
                 end if

                    Else
                   ' response.write sql
                        response.write ofv.ErrorDB(CNXM, volver)
                    End If

			end sub

%>
