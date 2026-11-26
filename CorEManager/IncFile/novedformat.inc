<%

Function EsqnomP(xtipo, xcodigo, ofv)

    Set xcn = ofv.conectar(ofv.strconn2)


    If xtipo = "RSK" Then

        StrSql = "SELECT * FROM  ORRSKCIRCUITO "
        StrSql = StrSql & " Where OBJID = 0" & xcodigo
        Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
        If Not xrs.EOF Then
            If xrs("APLIC") = "OR" Then
                xdescri = "CIRCUITO AUTOEVALUACION"
            Else
                xdescri = "PLAN DE CONTINGENCIAS"
            End If

            StrSql = "Select * From objetos where codobjeto = 0" & xrs("CODOBJETO")
            Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
            If Not xrs.EOF Then
                xiden = xrs(1)
                xobjdes = xrs(3)
            End If


        End If
        Call ofv.cerrarconsulta(xrs)

   '''''sTipobjetoX = Capitalize(xdescri) & ": " & Capitalize(xobjdes)
        sTipobjetoX = xdescri & ": " & xobjdes

    Else


        sTipobjetoX = xcodigo

        StrSql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
        Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
        If Not RSM.EOF Then




            If SESSION("APLMDCLI") = "BC" Then
            sTipobjeto1 = (RSM(1)) & ":&nbsp;"
            sTipobjeto2 = RSM(3) & ""
            sTipobjeto3 = " - " & (RSM(2) & "")
            sTipobjeto4 = " (" & (RSM(5) & "") & " "
            sTipobjeto5 = (RSM(6) & "") & ") "
            Else
        '''''sTipobjeto1 = Capitalize(RSM(1)) & ":&nbsp;"
             sTipobjeto1 = RSM(1) & ":&nbsp;"
        '''''sTipobjeto2 = Capitalize(RSM(3) & "")
             sTipobjeto2 = RSM(3) & ""

       '''''sTipobjeto3 = " - " & Capitalize(RSM(2) & "")
            sTipobjeto3 = " - " & RSM(2) & ""
       '''''sTipobjeto4 = " (" & Capitalize(RSM(5) & "")
            sTipobjeto4 = " (" & RSM(5) & " )"
            If UCase(RSM(5) & "") = "O" Then
             sTipobjeto5 = ")"
            Else
       '''''sTipobjeto5 =  " " & Capitalize(RSM(6) & "") & ") "
            sTipobjeto5 = " " & RSM(6) & "" & ") "
            End If
            End If
            If SESSION("FULLDOCN") = "1" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2
            ElseIf SESSION("FULLDOCN") = "2" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
            Else
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
            End If



        End If

    End If





    Call ofv.cerrarconn(xcn)


    EsqnomP = sTipobjetoX


End Function
Function noveformat(xmenu, XMENSEL,TNTabla)


    XTablaN = ""

    acctitblks = 0
    acctitlin = 0


    'xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "

    xsclase = " CLASS=ietx "
    xsclase2 = "class=istt "
    '''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "

    XTablaN = XTablaN & "<div id='DTNOV" & xmenu & "' style='width:100%; "


    If xmenu = 2 Then

        '    XTABLAN = XTablaN & " width:40%;"

    Else

        '   XTABLAN = XTablaN & " width:30%;"

    End If

    '********* ENCABEZADO DE SECCIONES
    XTablaN = XTablaN & "border-style:solid;border-width:3;border-color:#ffffff;' >"
    XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1     >"
    XTablaN = XTablaN & "<tr HEIGHT=25  >"
''''XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(Capitalize(atit(xmenu)) & "", " ", "&nbsp;") &  "</TD>"
    XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace( "", " ", "&nbsp;") &  " (menu " & xmenu & ")</TD>"
    XTablaN = XTablaN & "</TR>"
    XTablaN = XTablaN & "</table>"

    XTablaN = XTablaN & "<div id='DTNOVINT" & xmenu & "' style='"

    XTablaN = XTablaN & "border-style:solid;border-width:1;border-color:#ffffff;' >"

    XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=0  >"


    If TNTabla = 0 Then


        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;"  & "&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"

        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"


    Else

        For INTabla = 1 To TNTabla

            '    acctitlin  = acctitlin + 2


            xtmcod = ANTabla(INTabla, 0)
            xtipom = ANTabla(INTabla, 1)
            xtpri = ANTabla(INTabla, 2)
            xtfvig = ANTabla(INTabla, 3)
            xthvig = ANTabla(INTabla, 4)
            xtsadj = ANTabla(INTabla, 5)
            xtslnk = ANTabla(INTabla, 6)
            xfromusr = ANTabla(INTabla, 7)
            xtmsg = ANTabla(INTabla, 8)
            xtmsgimg = ANTabla(INTabla, 9)
            xtmsglnk = ANTabla(INTabla, 10)
            xtmsgTYPE = ANTabla(INTabla, 11)



            xfclas = "dt"
            sImage = "&nbsp;"

            Select Case xtpri
                Case 0

                    xfclas = "btbn"
                    sglmod = 5
                    sgltit = "Prioridad Alta."

                Case 1

                    xfclas = "btba"
                    sglmod = 4
                    sgltit = "Prioridad Normal."

                Case 2

                    xfclas = "btbs"
                    sglmod = 6
                    sgltit = "Prioridad baja."


            End Select

            sglmod = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
            sImage = ofv.GenLink("", xfclas & " title='" & sgltit & "' ", "", "", "", "", sglmod, "si")

            If xtfvig = dFechai Then
                XMARQ = sglmar
            Else
                XMARQ = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
            End If

            If xtipom = 5 Then
                xtfechavig = xthvig
                XTCOUNT = InStr(xtfechavig, " ")
                If XTCOUNT > 0 Then
                    xtfechavig = Left(xtfechavig, XTCOUNT)
                    xtfechavig = GTFechadorsh(CLng(CDate(xtfechavig)) - CLng(CDate("01-01-1900")))
                End If
            Else
                xtfechavig = GTFechadorsh(xtfvig)
                XTCOUNT = InStr(xtfechavig, " ")
                If XTCOUNT > 0 Then xtfechavig = Left(xtfechavig, XTCOUNT)


            End If

            Xnlnk = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
            Xnadj = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

            If xtsadj = "S" Then Xnadj = sgladj
            If xtslnk = "S" Then Xnlnk = sgllnk









            xtmsgVRS = "0" & ANTabla(INTabla, 12)
            xtmsgVRS = CInt(xtmsgVRS)
            xtmsgTIT = ANTabla(INTabla, 13)
            xtmsgCOP = ANTabla(INTabla, 14)
            xtmsgimg = ANTabla(INTabla, 15)
            xtmsgVID = ANTabla(INTabla, 16)


            xtmsgVAENC = ANTabla(INTabla, 17)
            xtmsgVAFORO = ANTabla(INTabla, 18)
            xtmsgENCCOD = ANTabla(INTabla, 19)
            xtmsgFOROCOD = ANTabla(INTabla, 20)

            xtmsgdoccod = ANTabla(INTabla, 21)


            If xtmsgVRS = 2 Then

                acctitlin = acctitlin + 2

                'xtmsglnk    = replace(xtmsglnk,"XXMSGXX","<font STYLE='color:#990000;font-size:10pt;font-weight:bold;' >" & xtmsgTIT & "</font>")
                xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

                '  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                '  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                '  XTABLAN = XTablaN &  "</TR>"

                XTablaN = XTablaN & "<tr HEIGHT='20' >"
                XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
                For yy = 1 To 6
                    XTablaN = XTablaN & "<td class=UT width='15%'  >&nbsp;</TD>"
                Next
                XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"

                ''  TIPO DE NOVEDAD NOTICIA DEL DIA
                If xtipom = "1" Then

                    xSolasp = ""
                    xSolasp = xSolasp & "novedmenuview.asp?xxmsg=" & CStr(xtmcod) & "&Pantalla=MAIN"




                    XTablaN = XTablaN & "<form name='VNGP" & CStr(xtmcod) & "' action='" & xSolasp & "' "
                    XTablaN = XTablaN & " method=post target='fesquemas' style='display:none;' >"
                    ' XTablaN = XTablaN & " method=post target='_new' style='display:none;' >"
                    XTablaN = XTablaN & "</form>"



                    xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

                    XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"


                     If SESSION("APLMDCLI") = "BGAL" Then
                     XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8   ><MARQUEE>" & xtmsglnk & "</MARQUE></a></TD>"

                     Else

                    XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8   >" & xtmsglnk & "</a></TD>"


                    End If


                    XTablaN = XTablaN & "</TR>"





                    '' Otros tipos de novedad
                Else



                    XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"
                    XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8  >" & xtmsglnk & "</TD>"

                    XTablaN = XTablaN & "</TR>"

                End If


                If xtipom = "9" Then

                    acctitlin = acctitlin + 2

                    XTMSGSEND = ""
                    XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP

                    XTablaN = XTablaN & "<tr  >"
                    XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B>"
                    XTablaN = XTablaN & "<font color='#AA0000' size=2 >&#0187;</font>"
                    XTablaN = XTablaN & "<font color='#0000AA' size=1 >&#0124;</font>"
                    XTablaN = XTablaN & "</B></TD>"
                    XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
                    XTablaN = XTablaN & XTMSGSEND & "</TD>"
                    XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"


                    XTMSGSEND = ""
                    XTMSGSEND = XTMSGSEND & "<font color=#888888 style='font-weight:bold;' size=1 >recomendado por " & ANTabla(INTabla, 7) & "</font>"

                    XTablaN = XTablaN & "<tr HEIGHT='20' >"
                    XTablaN = XTablaN & "<td class=UT align='right' colspan=8 STYLE='background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"

                End If


                '''''''''''''''''''''''''''''
                XTablaN = XTablaN & "<tr HEIGHT='20' >"
                XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"


                If xtmsgdoccod <> "0" Then



                    StrSql = "Select * From objinfo where codobjeto = 0" & xtmsgdoccod & " "
                 '' response.write strsql
                    Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs1.EOF Then
'aca esta el error

                 ''''' AGREGADO PARA MEJORAS VERSION 4.6
                       xtmsgCOPP = rs1(6)
                       if xtmsgCOPP <> "" then
				          xtmsgCOP = xtmsgCOP & "<br><br>" & rs1(6)
                       end if
                 ''''' AGREGADO PARA MEJORAS VERSION 4.6 FIN

                       xtmsgCOP = xtmsgCOP & "<br><br>" & rs1(4) '& xtmsgdoccod
                       xtmsgimg = rs1(7)
                       xtmsgVID = rs1(9)

                    End If

                    StrSql = "Select * From objinfoIMG  where codobjeto = 0" & xtmsgdoccod & " ORDER BY objSECUEN, OBJID"

                    wspimgdat = 6

                Else

                    StrSql = "Select * From NOVEDADESIMG  where NOVEDAD = '" & xtmcod & "' ORDER BY SECUEN, OBJID"
                    wspimgdat = 5
                End If



                If (xtmsgimg & "") = "1" Then

                    '   response.write Strsql

                    Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs1.EOF Then



                        XTablaN = XTablaN & "<tr bgcolor='#F0F0F0' >"
                        XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

                        XTablaN = XTablaN & "<table cellspacing=0 cellpadding=1  width=100% >"
                        XTablaN = XTablaN & "<tr   >"


                        XSFSTYLE = " STYLE='display:block;' "

                        Do While Not rs1.EOF

                            acctitlin = acctitlin + 10

                            xCelda = ""
                            xCelda = xCelda & "<div id='DIMGN" & INTabla & "' style='margin-top:1;margin-left:1;"
                            xCelda = xCelda & "width:90%;overflow:auto;"
                            xCelda = xCelda & ""
                            xCelda = xCelda & "' >" & Chr(13)
                            xCelda = xCelda & "<IMG style='border-style:solid;border-width:1;border-color:#888888;' SRC=" & Chr(34) & rs1(wspimgdat) & Chr(34) & ">"
                            xCelda = xCelda & "</div>" & Chr(13)

                            XTablaN = XTablaN & "<td CLASS=UT align='center' " & XSFSTYLE & " width=100% >" & xCelda & "</TD>"
                            XSFSTYLE = " STYLE='display:none;' "


                            rs1.movenext
                        Loop

                        XTablaN = XTablaN & "</TR>"
                        XTablaN = XTablaN & "</table>"

                        XTablaN = XTablaN & "</TD>"
                        XTablaN = XTablaN & "</TR>"


                        XTablaN = XTablaN & "<tr HEIGHT='20' >"
                        XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
                        XTablaN = XTablaN & "</TR>"


                    End If






                ElseIf (xtmsgimg & "") = "2" Then

                    acctitlin = acctitlin + 10


                    XTablaN = XTablaN & "<tr bgcolor='#F0F0F0' >"
                    XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
                    XTablaN = XTablaN & "<a class=istt target='_blank' href=" & xtmsgVID & "> " & xtmsgVID & "</a>"
              '  XTablaN = XTablaN & xtmsgVID
                    '   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
                    XTablaN = XTablaN & "</TD>"
                    XTablaN = XTablaN & "</TR>"


                    XTablaN = XTablaN & "<tr HEIGHT='20' >"
                    XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"

                End If

                acctitlin = acctitlin + 2

                '**** COPETE NOTICIA DEL DIA
                XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
                XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP
           '''''response.write xtmsgCOP & " xtmsgCOP - mensaje "
                XTablaN = XTablaN & "<tr  >"
                XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B><font color='#AA0000' size=2 >&#0149;</font></B></TD>"
                XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='text-align:justify;' >"
                ' LINK VER MAS
                If xtipom = "1" Then
                    url = "novedmenuview.asp?xxmsg=" & xtmcod
                    sFuncion = " onclick=" & Chr(34) & "document.all.VNGP" & CStr(xtmcod) & ".submit();WKlevanta('" & CStr(xtmcod) & "');" & Chr(34)
                    xtmsglnk = ofv.GenLink("", " istt style='cursor:pointer ;' title=' Ver Mas ...' ", "", "", "", sFuncion, "Ver Mas", "si")
                    XTablaN = XTablaN & XTMSGSEND & "<br> <br>" & xtmsglnk
                Else
                XTablaN = XTablaN & XTMSGSEND & "</TD>"
                End If
               ' XTablaN = XTablaN & "</TD>"







                ''' ********
                XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"




                ''''' PREGUNTA SI VA ENCUESTA '''''''''''''''''''''''

                '    xtmsgVAENC      = ANTabla(INTabla,17)
                '    xtmsgVAFORO     = ANTabla(INTabla,18)
                '    xtmsgENCCOD     = ANTabla(INTabla,19)
                '    xtmsgFOROCOD    = ANTabla(INTabla,20)



                If xtmsgVAENC = "S" Then

                    acctitlin = acctitlin + 2

                    StrSql = "Select * From emencuestas  where objid = 0" & xtmsgENCCOD
                    Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs1.EOF Then

                 '''''''xtmsgTITENC = Capitalize(rs1(17) & "")
                        xtmsgTITENC = rs1(17) & ""
                        xtmsgENCDER = "S"


                        If (rs1(2) & "") = "1" Then
                            StrSql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & xtmsgENCCOD
                            StrSql = StrSql & " Order By BLKSECUEN,objid DESC"
                            Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rs1.EOF Then
                                sblktvoto = rs1(4) & ""
                                sblkcodigo = rs1(0) & ""
                                If sblktvoto = "2" Then
                                    xtmsgENCDER = "N"
                              ''''''xtmsgTITENC = Capitalize(rs1(3) & "")
                                    xtmsgTITENC = rs1(3) & ""



                                    XTablaN = XTablaN & "<tr  >"
                                    XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' >&nbsp;</TD>"
                                    XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='color:#000000;' align='center' ><br>"

                                    XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=1  style='background-color:#eeeeee;border-style:solid;border-width:1;border-color:#dddddd;' >"

                                    XTablaN = XTablaN & "<tr HEIGHT=20  >"
                                    XTablaN = XTablaN & "<td class=vtI COLSPAN=3  >Encuesta</TD>"
                                    XTablaN = XTablaN & "</TR>"


                                    XTablaN = XTablaN & "<tr HEIGHT=20  >"
                                    XTablaN = XTablaN & "<td class=utI width=85% style='font-size:8pt;font-weight:normal;' >" & xtmsgTITENC & "</TD>"

                                    WSPDIVN = "DENC" & xtmsgENCCOD

                                    StrSql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & xtmsgENCCOD
                                    StrSql2 = StrSql2 & " and BLKcodigo = 0" & sblkcodigo
                                    StrSql2 = StrSql2 & " Order By ITMSECUEN,objid DESC"

                                    Set rs2 = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)
                                    Do While (Not rs2.EOF)

                                        acctitlin = acctitlin + 2


                                        WSPFRMN = "FENC" & xtmsgENCCOD & "IT" & rs2(0)
                                        WSPFRMA = "EMENCUESTASACT.ASP?TP=ES&OP=ACT&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
                                        WSPFRMR = "EMENCUESTASACT.ASP?TP=ES&OP=RES&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
                                        XTablaN = XTablaN & "<td class=ut width=5% align='center' >"
                                        WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"
                                        XTablaN = XTablaN & "<input type='button' style='width:45px;background-color:#cc0000;color:#ffffff;font-weight:bold;border-style:solid;border-width:1;border-color:#888888;' "
                                        XTablaN = XTablaN & "  value='&nbsp;" & rs2(4) & "&nbsp;' "
                                        XTablaN = XTablaN & "  onclick='"
                                        '  XTABLAN = XTablaN &  "alert("" voto:" & rs2(4) & " "");"
                                        '  XTABLAN = XTablaN &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
                                        XTablaN = XTablaN & "document." & WSPFRMN & ".submit();"
                                        XTablaN = XTablaN & "document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";"
                                        XTablaN = XTablaN & "' >"
                                        XTablaN = XTablaN & "</TD>" & WSPFORME
                                        rs2.movenext
                                    Loop

                                    acctitlin = acctitlin + 2

                                    xCelda = ""
                                    xCelda = xCelda & "<div id='" & WSPDIVN & "' style='z-index:1;margin-top:1;margin-left:1;display:none;"
                                    xCelda = xCelda & "width:99%;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
                                    xCelda = xCelda & "background-color:#F9F9F9;"
                                    xCelda = xCelda & "' >" & Chr(13)
                                    xCelda = xCelda & "&nbsp;"
                                    xCelda = xCelda & "</div>" & Chr(13)


                                    XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"
                                    XTablaN = XTablaN & "</TR>"

                                    XTablaN = XTablaN & "<tr HEIGHT=20  >"
                                    XTablaN = XTablaN & "<td class=utI width=85%  >&nbsp;</TD>"
                                    XTablaN = XTablaN & "<td class=vt COLSPAN=2 ALIGN='CENTER' ><A CLASS=NT HREF='" & WSPFRMR & "' TARGET='fesquemas' "
                                    XTablaN = XTablaN & " ONCLICK='document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";' "
                                    XTablaN = XTablaN & "><B>Resultados</B></TD>"
                                    XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"

                                    XTablaN = XTablaN & "</TR>"

                                    XTablaN = XTablaN & "<tr   >"
                                    XTablaN = XTablaN & "<td class=vtI COLSPAN=4 >&nbsp;</TD>"
                                    XTablaN = XTablaN & "</TR>"



                                    XTablaN = XTablaN & "</table>" & xCelda
                                    XTablaN = XTablaN & "</TD>"
                                    XTablaN = XTablaN & "<td class=ut   >&nbsp;</TD>"
                                    XTablaN = XTablaN & "</TR>"





                                End If

                            End If


                        End If





                    End If

                End If


                ''''''''''''''''''''''''''''

                XTablaN = XTablaN & "<tr HEIGHT='20' >"
                XTablaN = XTablaN & "<td class=UT colspan=8 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"











            Else

                acctitlin = acctitlin + 1


                ''''  xtmsg    = replace(ltrim(xtmsg & "")," ","&nbsp;")
                xtmsg = LTrim(xtmsg & "")
                xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsg & "")


                If xtmcod = "0" Then

                    XTMSGSEND = "<font color=#000000 >" & xtmsglnk & "</font>"

                Else

                    XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
                    XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                    XTMSGSEND = XTMSGSEND & "<font color=#000000 >&nbsp;" & xtmsglnk & "</font>"

                End If



                XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " HEIGHT='25' >"
                XTablaN = XTablaN & "<td class=" & xfclas & " align='center' >" & xtmsgimg & "</TD>"
                XTablaN = XTablaN & "<td class=uti colspan=6 STYLE='text-align:justify;' >" & XTMSGSEND & "</TD>"
                XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"


                If xtmcod <> "0" Then

                    ' acctitlin  = acctitlin + 1


                    XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=5 >"
                    XTablaN = XTablaN & "<td class=ut colspan=8 valign=top align=right style='border-top-style:solid;border-top-width:1;border-top-color:#eeeeee;' >&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"

                End If


            End If


            If xTrColor = "#FFFFFF" Then
                xTrColor = "#eeeeff"
            Else
                xTrColor = "#FFFFFF"
            End If


            '  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
            '  XTABLAN = XTablaN &  "<td class=UT colspan=8  ><HR  size=1 style='width:400;color:#ff0000;' >&nbsp;</TD>"
            '  XTABLAN = XTablaN &  "</TR>"




        Next


    End If

    XTablaN = XTablaN & "<tr HEIGHT='20' >"
    XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
    XTablaN = XTablaN & "</TR>"



    XTablaN = XTablaN & "</table>"
    XTablaN = XTablaN & "<HR  size=1 style='width:350;color:#eeeeee;' >"
    XTablaN = XTablaN & "</DIV>"

    XTablaN = XTablaN & "</DIV>"


'response.write TNTabla
    If TNTabla = 0 Then
        noveformat = ""
    Else

        If xmenu = 4 Then atmensajes = TNTabla

        ttitsecc = ttitsecc + 1

   '''''atitsecciones(ttitsecc, 0) = Capitalize(atit(xmenu))


        noveformat = XTablaN
    End If


End Function
%>