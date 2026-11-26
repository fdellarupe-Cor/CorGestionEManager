
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok = ofv.CheckUsuario()
If ok <> "Y" Then
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    shtml = ok
Else
    Dim AposLevel(8)
    XENTRO1 = True
    XENTRO2 = True
    XENTRO3 = True
    XENTRO4 = True
    XENTRO5 = True
    XENTRO6 = True

    aplcheck()
    colorPuesto="#77AADD"
    SESSION("IDRET")=""
    SESSION("BSQ")=""
    ' for i = 0 to 8
    ' AposLevel(i)="A"
    ' next
    ' Amaxlevel="0,0,"
    Generar = ""
    NIVELCONT = 90000
    'response.buffer=false
    Set cn = ofv.conectar(ofv.strconn2)
    Set cn2 = ofv.conectar(ofv.strconn0)
    Set cn3 = ofv.conectar(ofv.strconn0)


    StrSql = "select *  from listadepuestos where  site='" & SESSION("SUN") & "' and tipo='G'"
    ' strsql="select puestodepen from relapuesto where puestodepen>0 group by  puestodepen having  count(*)  >1"
    'response.write strsql
    Set rsasist = ofv.crearconsultaEx(StrSql, cn, 1, "")
    idsgen = "0"
    If Not rsasist.EOF Then
        objiAsis = rsasist(0)

        Do While Not rsasist.EOF
            idsgen = idsgen & "," & rsasist(0)
            rsasist.movenext

        Loop
    Else
        objiAsis = -1


    End If
    Call ofv.cerrarconsulta(rsasist)

    If request.querystring("mseg") <> "" Then
        session("mseg") = request.querystring("mseg")
    End If

    nID = ""

    sSector = ""
    nSeg = ""

    nID = request.querystring("ID")

    If session("MSEG") = "E" Then

        nidu = session("SGU_idcostcenter")  'rs(0)
        NIVEL = session("SGU_idnivel")    'rs(1)

        xxx = getinfouser(session("usuario"))

        If nID = "" Then
            nID = nidu
        End If

    Else
        NIVEL = 0

        sSector = request.querystring("Sector")
        sSector = ofv.ConvertirCar(sSector, "_", " ")
        sCodin = "&nbsp;"
        If nID = "" And request.Form.Count > 0 Then
            nID = request.Form(1)
        End If
    End If

    nID = "0" & Replace(nID, "nsnc", "")

   ' response.write nid & "---"
    If Not nID > "0" Then
        shtml =  ofv.MenuHeader("", "")
        'shtml=replace(shtml,"</HEAD>","222</HEAD>")
        shtml = shtml & "<br><br><br><br><br><center><p><b><font face=verdana size=4 color=#993333>DEBE SELECCIONAR UN AREA O SECTOR<br>PARA GRAFICAR.</FONT></P>" & Chr(13)
        shtml = shtml & "<br></body></html>"
    Else
        sSECSITE = 1
        sSECNIVEL = 0
        SNIVELA = "S/D"
        ndotacion = 1
        nivelj0 = 0
        If nID <> "" Then
            'sSQL = "Select descri,codigo,nivelj,dotacion,continuador From listadepuestos Where objid = " & nID
            sSQL = "Select l.descri,l.codigo,l.nivelj,l.dotacion,l.continuador ,p.color From listadepuestos l "
            ssql =sSql & " inner join nivelpuesto p on (p.nivelp=l.niveljREAL and l.site='" & SESSION("SUN") & "') Where l.objid = " & nID
          ' RESPONSE.WRITE SSQL & "<BR>"
          ''   XXXX = GETOBJALOGUPD("ACCMOD")

              XXXX = GETFECH()


              XLFECDAT2  = XLFECDAT
              XLHORDAT2  = XLHORDAT

            Set rs = ofv.crearconsultaEx(sSQL, cn, 1, "")
            If Not rs.EOF Then
                sSector = rs(0)
                scontinuador = rs("continuador")
                ndotacion = rs("dotacion")
                sCodin = rs(1)
                sSECSITE = 1
                sSECNIVEL = request.Form("SECNIVEL")
                nivelj0 = rs(2)
                colorPuesto=rs("color")
                If sSECSITE <> "" Then
                    session("SSITE") = sSECSITE
                Else
                    sSECSITE = session("SSITE")
                End If

                If sSECNIVEL = "" Then sSECNIVEL = 0



                sSUCA = "1"
                sSUCN = "1"



                sSQL = "SELECT * FROM NIVELPUESTO where NIVELP = 0" & rs(2)
                Set RS1 = ofv.crearconsulta2(sSQL, cn)
                If Not RS1.EOF Then
                    SNIVELA = RS1(1)
                    '''  response.write "NIV " & SNIVELA
                End If

            End If
            Call ofv.cerrarconsulta(rs)
        End If


        ' response.write NID & "----" & sSector
        sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=2>&#0128;</FONT></B>"
        sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"
        sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

        If nID <> "" Then




            shtml = ofv.MenuHeader("", "")

'            shtml=replace(shtml,"</HEAD>","<meta http-equiv=""X-UA-Compatible"" content=""IE=7""/></HEAD>")

            shtml = shtml & Chr(13) & "<script type='text/javascript' src='incfile/ECOTreeN.js' ></script><link type='text/css' rel='stylesheet' href='incfile/ECOTree.css' />"
            shtml=  shtml & "<script type='text/javascript' src='JSGRAF/jspdf.js'></script>"


			shtml=  shtml & "<script type='text/javascript' src='JSGRAF/jspdf.plugin.addimage.js'></script> "
			shtml=  shtml & " <script type='text/javascript' src='JSGRAF/jspdf.source.js'></script>"

    '        shtml = shtml & Chr(13) & "<style>v\: * { behavior:url(#default#VML); display:inline-block }</style>"

    '        shtml = shtml & Chr(13) & "<style>"
    '        shtml = shtml & Chr(13) & " .copy {"
    '        shtml = shtml & Chr(13) & " font-family : 'Verdana';"
    '        shtml = shtml & Chr(13) & " font-size : 9px;"
    '        shtml = shtml & Chr(13) & " color : #000000;"
    '        shtml = shtml & Chr(13) & " } "
    '        shtml = shtml & Chr(13) & "</style>"
            shtml = shtml & "</head>"



            xin = 0
            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)
            ReDim Preserve vecsecn(xin)
            ReDim Preserve vecALT(xin)

            'Sector Principal

            wwwsettree = ""
            ''   wwwsettreeAN = "400"
            ''   wwwsettreeAL = "200"


      ''      wwwsettreeAN = "290"
       ''     wwwsettreeAL = "150"


            wwwsettreeAN = "190"
            wwwsettreeAL = "50"




            sSQL = "SELECT L_1.objid,L_1.descri,L_1.codigo,L_1.tipo,L_1.segmento,L_1.nivelj,L_1.nivelj,'1','1',LR.puestoREL"
            sSQL = sSQL & " FROM listadepuestos L , RelaPuesto LR, listadepuestos L_1 "

            sSQL = sSQL & " where l.site='" & SESSION("SUN") & "' and L.objid = LR.puestoDEPEN and lR.puestoMASTER = L_1.objid and L.objid = " & nID

             'response.write " 1   " & SSQL & "<br><br>" & chr(13)
         '  colorPuesto="#77AADD"
            Set rs = ofv.crearconsulta2(sSQL, cn)
            If rs.EOF Then
                vecsec(xin) = 0
                vecsecd(xin) = sSector
                vecsecn(xin) = 0
                vecALT(xin) = 0
                'nivelj0=0
                xsegmento0 = "1"
            Else

                SCODREL = rs(9)

                vecsec(xin) = CLng(rs(0))
                vecsecd(xin) = sSector
                vecsecn(xin) = rs(5)
                vecALT(xin) = 1
                xsegmento0 = "0" & rs(4)
              '  colorPuesto=rs("color")
                '''''''''''''
                '  nivelj0=rs(5)

            End If

           ' response.write colorPuesto 'session("mseg") & "-" & nid & "-" & nidu & "-" & (cstr(nid) = cstr(nidu))

            If session("mseg") = "E" And CStr(nID) = CStr(nidu) Then vecsec(xin) = 0



            If vecsec(0) <> 0 Then
                ' vecsec(0)=vecsec(0) 'buscapadre(vecsec(0))
                shref = "OrganigramaPuestosgraf.asp?ID=" & vecsec(0) & "&Sector=" & ofv.ConvertirCar(vecsecd(0), " ", "_")

                sinput = "<input type=button value='Subir' style='color:#000000;font-weight:bold;background-color:#ccddee;' onclick='document.OPSUBE.submit();'  ><form id='OPSUBE' name='OPSUBE' action='" & shref & "' method='post' style='display:none;' >"
                slink = ofv.GenLink("", "nt", shref, "", "", "", sinput, "si")


                wwwsettree = wwwsettree & "myTree.add2(0,-1,'SUBIR',50,30,'" & session("CMFMENU") &"','#FFFFFF','" & shref & "',null,'1','NO');"
               ' wwwsettree = wwwsettree & "myTree.add2(0,-1,'SUBIR',50,30,'#77AADD','#000000','" & shref & "',null,'1','NO');"

                shrefold = "OrganigramaPuestosgraf.asp?ID=" & vecsec(0) & "&Sector=" & ofv.ConvertirCar(vecsecd(0), " ", "_")


                '   shref = "OrganigramaPuestosgraf.asp?ID=" & nid & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")

                '    shref = "organigramaUNFunc.asp?ID=" & nid & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")


                shref = "GENPUESTOFUNCION.asp?PUESTO=" & nID & "&NOMSECTOR=" & ofv.ConvertirCar(vecsecd(0), " ", "_")
                sinput = "<font face=verdana color=#000080 size=1><b>" & vecsecd(0) & "</b></font>"
                slink = ofv.GenLink("", "nt", shref, "", "", "", sinput, "si")

                wwwsettree = wwwsettree & "myTree.add2(1,0,'" &  (vecsecd(0) ) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"&colorpuesto&"','#000000','" & shref & "',null,'"&ndotacion&"','" & scontinuador &"','0');"


            Else


                shref = "GENPUESTOFUNCION.asp?PUESTO=" & nID & "&NOMSECTOR=" & ofv.ConvertirCar(vecsecd(0), " ", "_")
                sinput = "<font face=verdana color=#993333 size=1><b>" & vecsecd(0) & "</b></font>"
                slink = ofv.GenLink("", "nt", shref, "", "", "", sinput, "si")

                wwwsettree = wwwsettree & "myTree.add2(1,-1,'" & (vecsecd(0)) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"&colorpuesto	&"','#000000','" & shref & "',null,'1','NO','0');"

            End If




            sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
            sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, ECO_1.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR,ER.POSY,eco_1.tipo,nj.color "
            sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
            sSQL = sSQL & " where e.site='" & SESSION("SUN") & "' and E.objid = ER.puestoMASTER and "
            sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
            sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
            sSQL = sSQL & " eco_1.niveljREAL = NJ.NIVELP  AND "

            sSQL = sSQL & " E.objid = " & nID & "   "
            'sSQL = sSQL & " and (segmento.codseg =  '" & xsegmento0 & "' OR ECO_1.TIPO='GA') "
            sSQL = sSQL & " order by  ECO_1.SECUEN  ,ECO_1.nivelj  "



           '  response.write " 2   " & SSQL & "<br><br>" & chr(13)
            Set rs = ofv.crearconsulta2(sSQL, cn)

            'Texto de Sectores Continuadores y Lineas
            Do Until rs.EOF
                ''''''''''''''''
                SCODREL = rs(9)

                xin = xin + 1
                ReDim Preserve vecsec(xin)
                ReDim Preserve vecsecd(xin)
                ReDim Preserve vecsecn(xin)
                ReDim Preserve vecALT(xin)
                vecsec(xin) = CLng(rs(0))
                vecsecd(xin) = rs(1)
                vecsecn(xin) = rs(5)
                vecALT(xin) = 2
                nivelj1 = rs(5)
                xsegmentoSALTAR = CLng("0" & rs("POSY"))

                xsegmento = "0" & rs(4)

                xgrafica = "" & rs(13)

                If xgrafica = "S" Then
                    shref = "OrganigramaPuestosgraf.asp?ID=" & vecsec(xin) & "&Sector=" & ofv.ConvertirCar(vecsecd(xin), " ", "_")
                Else
                    shref = "GENPUESTOFUNCION.asp?PUESTO=" & vecsec(xin) & "&NOMSECTOR=" & ofv.ConvertirCar(vecsecd(xin), " ", "_")
                End If

                pad = 1
                xsegmentoSALTAR = (nivelj1 - nivelj0) - 1
                'response.write nivelj1 & "---" & nivelj0 & "<br>"
                If xsegmentoSALTAR > 0 Then
                    pad = AGREGARNIVEL(xsegmentoSALTAR, pad)
                    '     vecALT(xin)=vecALT(xin) +  xsegmentoSALTAR

                End If
                If rs("tipo") = "S" Then

                    wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & -1 & ",'" & rs(1) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"& RS("COLOR")&"','#000000','" & shref & "',null,'" & rs("DOTACION") & "','" & rs("CONTINUADOR") & "','G');"
                ElseIf rs("tipo") = "G" Then
                    objiAsis = rs(0)
                    wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  rs(1) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"& RS("COLOR") &"','#000000','" & shref & "','"&objiAsis&"','"&RS("DOTACION")&"','SI','"&rs(0)&"');"
                Else
                    wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" & rs(1) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'" & rs("color")& "','#000000','" & shref & "',null,'" & rs("DOTACION") & "','" & rs("CONTINUADOR") & "','" & rs(0) & "');"
                End If
                xinw = xin + 1




                sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
                sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, ECO_1.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR, ER.POSY,eco_1.tipo,nj.color  "
                sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
                sSQL = sSQL & " where e.site='" & SESSION("SUN") & "' and E.objid = ER.puestoMASTER and "
                sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
                sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
                sSQL = sSQL & " eco_1.niveljREAL = NJ.NIVELP  AND "
                sSQL = sSQL & " E.objid =" & rs(0) & "   "
                sSQL = sSQL & " AND  ECO_1.TIPO <> 'S' and (segmento.codseg =  '" & xsegmento & "' OR ECO_1.TIPO='GA') "
                sSQL = sSQL & " order by ECO_1.SECUEN  ,ECO_1.nivelj"


                ' response.write " 3   " & SSQL & "<br><br>" & chr(13)
                Set rsn = ofv.crearconsulta2(sSQL, cn)

                'Texto de Sectores Continuadores y Lineas
                pp = 0
                zxwin = 0
                Do Until rsn.EOF
                    ''''''''''''''''
                    xSCODREL = rsn(9)
                    xin = xin + 1

                    xvecsec = CLng(rsn(0))
                    xvecsecd = rsn(1)
                    xvecsecn = rsn(5)
                    xsegmentoSALTAR = 2    'CLNG("0" & RSN("POSY"))
                    pp = xsegmentoSALTAR
                    xsegmento2 = "0" & rsn(4)
                    nivelj2 = rsn(5)
                    ReDim Preserve vecALT(xin)

                    xgrafica = "" & rsn(13)

                    If xgrafica = "S" Then
                        shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec & "&Sector=" & ofv.ConvertirCar(xvecsecd, " ", "_")
                    Else
                        shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec & "&NOMSECTOR=" & ofv.ConvertirCar(xvecsecd, " ", "_")
                    End If


                    pad = xinw
                    xsegmentoSALTAR = (nivelj2 - nivelj1) - 1
                    '    if pp <> xsegmentoSALTAR then
                    '     pp=xsegmentoSALTAR
                    '     zxwin=xinw
                    '    end if
                    '   response.write nivelj2 & "---" & nivelj1 & "---" & xinw &"---" & xsegmentoSALTAR  &"<br>"
                    If xsegmentoSALTAR > 0 Then
                        pad = AGREGARNIVEL(xsegmentoSALTAR, pad)
                        '    vecALT(xin)=vecALT(xin) +  xsegmentoSALTAR

                    End If

                    ' wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN("DOTACION")&"','"&RSN("CONTINUADOR")&"','"&rsn(0)&"');"
                    '        if clng(RSn(0))=objiAsis AND XENTRO1 THEN
                    '                XENTRO1=FALSE
                    '                wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN("DOTACION")&"','"&RSN("CONTINUADOR")&"','"&rsn(0)&"');"
                    '          ELSEIF  clng(RSn(0))=objiAsis  THEN
                    '       if clng(RSn(0))=objiAsis  THEN
                    If rsn("tipo") = "G" Then
                        objiAsis = rsn(0)
                        wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"&rsn("color")&"','#000000','" & shref & "','"&objiAsis&"','"&RSN("DOTACION")&"','SI','"&rsn(0)&"');"
                    Else
                        wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" & (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"&rsn("color")&"','#000000','" & shref & "',null,'" & rsn("DOTACION") & "','" & rsn("CONTINUADOR") & "','" & rsn(0) & "');"

                    End If



                    xinw2 = xin + 1

                    '''''''''''''''''''''''

                    '   sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
                    '   sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN "
                    '   sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento "
                    '   sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
                    '   sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
                    '   sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
                    '   sSQL = sSQL & " E.objid = " & rsn(0) & "   "
                    '   sSQL = sSQL & " and segmento.codseg =  '" & xsegmento2 & "' "
                    '   sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "


                    sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
                    sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, ECO_1.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR, ER.POSY,eco_1.tipo ,nj.color"
                    sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
                    sSQL = sSQL & " where e.site='" & SESSION("SUN") & "' and  E.objid = ER.puestoMASTER and "
                    sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
                    sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
                    sSQL = sSQL & " eco_1.niveljREAL = NJ.NIVELP  AND "
                    sSQL = sSQL & " E.objid =" & rsn(0) & "   "
                    sSQL = sSQL & " AND  ECO_1.TIPO <> 'S'  and (segmento.codseg =  '" & xsegmento2 & "' OR ECO_1.TIPO='GA') "
                    sSQL = sSQL & " order by ECO_1.SECUEN  ,ECO_1.nivelj "


                      '    response.write " 4   " & SSQL & "<br><br>" & chr(13)
                    Set rsn2 = ofv.crearconsulta2(sSQL, cn)

                    'Texto de Sectores Continuadores y Lineas
                    Do Until rsn2.EOF
                        ''''''''''''''''
                        xSCODREL2 = rsn2(9)

                        xin = xin + 1

                        xvecsec2 = CLng(rsn2(0))
                        xvecsecd2 = rsn2(1)
                        xvecsecn2 = rsn2(5)
                        xsegmentoSALTAR = CLng("0" & rsn2("POSY"))
                        xsegmento3 = "0" & rsn2(4)
                        nivelj3 = rsn2(5)
                        xgrafica = "" & rsn2(13)

                        If xgrafica = "S" Then
                            shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec2 & "&Sector=" & ofv.ConvertirCar(xvecsecd2, " ", "_")
                        Else
                            shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec2 & "&NOMSECTOR=" & ofv.ConvertirCar(xvecsecd2, " ", "_")
                        End If

                        pad = xinw2
                        xsegmentoSALTAR = (nivelj3 - nivelj2) - 1
                        '    response.write nivelj3 & "---" & nivelj2 & "<br>"
                        If xsegmentoSALTAR > 0 Then
                            pad = AGREGARNIVEL(xsegmentoSALTAR, pad)
                        End If
                        ' wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN2("DOTACION")&"','"&RSN2("CONTINUADOR")&"','"&rsn2(0)&"');"
                        '    if clng(RSn2(0))=objiAsis AND XENTRO2 THEN
                        '                        XENTRO2=FALSE
                        '                        wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN2("DOTACION")&"','"&RSN2("CONTINUADOR")&"','"&rsn2(0)&"');"
                        '                      ELSEIF  clng(RSn2(0))=objiAsis  THEN
                        ' if clng(RSn2(0))=objiAsis  THEN
                        If rsn2("tipo") = "G" Then
                            objiAsis = rsn2(0)
                            wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"& rsn2("color") & "','#000000','" & shref & "','"&objiAsis&"','"&RSN2("DOTACION")&"','SI','"&rsn2(0)&"');"
                        Else
                            wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" & (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'" & rsn2("color") &"','#000000','" & shref & "',null,'" & rsn2("DOTACION") & "','" & rsn2("CONTINUADOR") & "','" & rsn2(0) & "');"

                        End If




                        xinw3 = xin + 1

                        '''''''''''''''''''''''


                        sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
                        sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, ECO_1.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR, ER.POSY,eco_1.tipo ,nj.color"
                        sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
                        sSQL = sSQL & " where e.site='" & SESSION("SUN") & "' and E.objid = ER.puestoMASTER and "
                        sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
                        sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
                        sSQL = sSQL & " eco_1.niveljREAL = NJ.NIVELP  AND "
                        sSQL = sSQL & " E.objid =" & rsn2(0) & "   "
                        sSQL = sSQL & " AND  ECO_1.TIPO <> 'S' and (segmento.codseg =  '" & xsegmento3 & "' OR ECO_1.TIPO='GA') "
                        sSQL = sSQL & " order by ECO_1.SECUEN  ,ECO_1.nivelj "


                        ' response.write " 5   " & SSQL & "<br><br>" & chr(13)
                        Set rsn3 = ofv.crearconsulta2(sSQL, cn)

                        'Texto de Sectores Continuadores y Lineas
                        Do Until rsn3.EOF
                            ''''''''''''''''
                            xSCODREL3 = rsn3(9)

                            xin = xin + 1

                            xvecsec3 = CLng(rsn3(0))
                            xvecsecd3 = rsn3(1)
                            xvecsecn3 = rsn3(5)
                            xsegmentoSALTAR = CLng("0" & rsn3("POSY"))
                            xsegmento4 = "0" & rsn3(4)
                            nivelj4 = rsn3(5)
                            xgrafica = "" & rsn3(13)

                            If xgrafica = "S" Then
                                shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec3 & "&Sector=" & ofv.ConvertirCar(xvecsecd3, " ", "_")
                            Else
                                shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec3 & "&NOMSECTOR=" & ofv.ConvertirCar(xvecsecd3, " ", "_")
                            End If

                            pad = xinw3
                            xsegmentoSALTAR = (nivelj4 - nivelj3) - 1
                            If xsegmentoSALTAR > 0 Then
                                pad = AGREGARNIVEL(xsegmentoSALTAR, pad)
                            End If

                            ' if clng(RSn3(0))=objiAsis AND XENTRO3 THEN
                            '    XENTRO3=FALSE
                            '    wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd3) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN3("DOTACION")&"','"&RSN3("CONTINUADOR")&"','"&rsn3(0)&"');"
                            '  ELSEIF  clng(RSn3(0))=objiAsis  THEN
                            'if clng(RSn3(0))=objiAsis  THEN
                            If rsn3("tipo") = "G" Then
                                objiAsis = rsn3(0)
                                wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd3) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'" & rsn3("color") & "','#000000','" & shref & "','"&objiAsis&"','"&RSN3("DOTACION")&"','SI','"&rsn3(0)&"');"
                            Else
                                wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" & (xvecsecd3) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'" &rsn3("color") & "','#000000','" & shref & "',null,'" & rsn3("DOTACION") & "','" & rsn3("CONTINUADOR") & "','" & rsn3(0) & "');"

                            End If



                            ' otro nivel mas

                            sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
                            sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN,ECO_1.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR, ER.POSY,eco_1.tipo ,nj.color"
                            sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
                            sSQL = sSQL & " where e.site='" & SESSION("SUN") & "' and E.objid = ER.puestoMASTER and "
                            sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
                            sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
                            sSQL = sSQL & " eco_1.niveljREAL = NJ.NIVELP  AND "
                            sSQL = sSQL & " E.objid =" & rsn3(0) & "   "
                            sSQL = sSQL & " AND  ECO_1.TIPO <> 'S' and (segmento.codseg =  '" & xsegmento4 & "' OR ECO_1.TIPO='GA') "
                            sSQL = sSQL & " order by ECO_1.SECUEN  ,ECO_1.nivelj"
                            xinw4 = xin + 1

                            '       response.write " 5   " & SSQL & "<br><br>" & chr(13)
                            Set rsn4 = ofv.crearconsulta2(sSQL, cn)

                            'Texto de Sectores Continuadores y Lineas
                            Do Until rsn4.EOF
                                ''''''''''''''''
                                xSCODREL4 = rsn4(9)

                                xin = xin + 1

                                xvecsec4 = CLng(rsn4(0))
                                xvecsecd4 = rsn4(1)
                                xvecsecn4 = rsn4(5)
                                xsegmentoSALTAR = CLng("0" & rsn4("POSY"))
                                xgrafica = "" & rsn4(13)
                                xsegmento5 = "0" & rsn4(4)
                                nivelj5 = rsn4(5)



                                If xgrafica = "S" Then
                                    shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec4 & "&Sector=" & ofv.ConvertirCar(xvecsecd4, " ", "_")
                                Else
                                    shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec4 & "&NOMSECTOR=" & ofv.ConvertirCar(xvecsecd4, " ", "_")
                                End If

                                pad = xinw4
                                xsegmentoSALTAR = (nivelj5 - nivelj4) - 1


                                If xsegmentoSALTAR > 0 Then
                                    pad = AGREGARNIVEL(xsegmentoSALTAR, pad)
                                End If
                                '     if clng(RSn4(0))=objiAsis AND XENTRO4 THEN
                                '       XENTRO4=FALSE
                                '       wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd4) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN4("DOTACION")&"','"&RSN4("CONTINUADOR")&"','"&rsn4(0)&"');"
                                '     ELSEIF  clng(RSn4(0))=objiAsis  THEN
                                ' if clng(RSn4(0))=objiAsis  THEN
                                If rsn4("tipo") = "G" Then
                                    objiAsis = rsn4(0)
                                    wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd4) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"& rsn4("color") & "','#000000','" & shref & "','"&objiAsis&"','"&RSN4("DOTACION")&"','SI','"&rsn4(0)&"');"
                                Else
                                    wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" & (xvecsecd4) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"& rsn4("color") & "','#000000','" & shref & "',null,'" & rsn4("DOTACION") & "','" & rsn4("CONTINUADOR") & "','" & rsn4(0) & "');"

                                End If



                                ' otro nivel mas

                                ' otro nivel mas

                                sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
                                sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, ECO_1.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR, ER.POSY,eco_1.tipo ,nj.color"
                                sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
                                sSQL = sSQL & " where  e.site='" & SESSION("SUN") & "' and E.objid = ER.puestoMASTER and "
                                sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
                                sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
                                sSQL = sSQL & " eco_1.niveljREAL = NJ.NIVELP  AND "
                                sSQL = sSQL & " E.objid =" & rsn4(0) & "   "
                                sSQL = sSQL & " AND  ECO_1.TIPO <> 'S' and (segmento.codseg =  '" & xsegmento5 & "' OR ECO_1.TIPO='GA') "
                                sSQL = sSQL & " order by ECO_1.SECUEN  ,ECO_1.nivelj"
                                xinw5 = xin + 1

                                ' response.write " 7   " & SSQL & "<br><br>" & chr(13)
                                Set rsn5 = ofv.crearconsulta2(sSQL, cn)

                                'Texto de Sectores Continuadores y Lineas
                                Do Until rsn5.EOF
                                    ''''''''''''''''
                                    xSCODREL5 = rsn5(9)

                                    xin = xin + 1

                                    xvecsec5 = CLng(rsn5(0))
                                    xvecsecd5 = rsn5(1)
                                    xvecsecn5 = rsn5(5)
                                    xsegmentoSALTAR = CLng("0" & rsn5("POSY"))
                                    nivelj6 = rsn5(5)
                                    xgrafica = "" & rsn5(13)

                                    If xgrafica = "S" Then
                                        shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec5 & "&Sector=" & ofv.ConvertirCar(xvecsecd5, " ", "_")
                                    Else
                                        shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec5 & "&NOMSECTOR=" & ofv.ConvertirCar(xvecsecd5, " ", "_")
                                    End If

                                    pad = xinw5
                                    xsegmentoSALTAR = (nivelj6 - nivelj5) - 1
                                    If xsegmentoSALTAR > 0 Then
                                        pad = AGREGARNIVEL(xsegmentoSALTAR, pad)
                                    End If
                                    '          if clng(RSn5(0))=objiAsis AND XENTRO5 THEN
                                    '            XENTRO5=FALSE
                                    '            wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd5) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN5("DOTACION")&"','"&RSN5("CONTINUADOR")&"','"&rsn5(0)&"');"
                                    '          ELSEIF  clng(RSn5(0))=objiAsis  THEN

                                    'if clng(RSn5(0))=objiAsis  THEN
                                    If rsn5("tipo") = "G" Then
                                        objiAsis = rsn5(0)
                                        wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd5) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"& rsn5("color") & "','#000000','" & shref & "','"&objiAsis&"','"&RSN5("DOTACION")&"','SI','"&rsn5(0)&"');"
                                    Else
                                        wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" & (xvecsecd5) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'"& rsn5("color") & "','#000000','" & shref & "',null,'" & rsn5("DOTACION") & "','" & rsn5("CONTINUADOR") & "','" & rsn5(0) & "');"

                                    End If



                                    ' otro nivel mas


                                    rsn5.movenext
                                Loop


                                rsn4.movenext
                            Loop


                            rsn3.movenext
                        Loop



                        ''''''''''''''''''''



                        rsn2.movenext
                    Loop



                    ''''''''''''''''''''


                    rsn.movenext
                Loop



                ''''''''''''''''''''

                rs.movenext
            Loop
            Call ofv.cerrarconsulta(rs)



            xinS = 0
            ReDim Preserve vecsecS(xinS)
            ReDim Preserve vecsecSD(xinS)
            ReDim Preserve vecsecSN(xinS)



            sft = "<font face=verdana color=#000077 size=1>&nbsp;"
            sft2 = "<font face=verdana color=#770000 size=1><b>&nbsp;"
            sft3 = "<font face=verdana color=#ff0000 size=1>&nbsp;"
            sft1 = "<font face=verdana color=#770000 size=1>&nbsp;"
            sff = "</font>"




             BTIT1 = "<A CLASS=CC HREF='BSQORGANIGRAMA.ASP?BSQ=PORDEPENDENCIA&nid=" & nID & "'>BUSCAR&nbsp;POR&nbsp;DEPENDENCIA</A>"
             BTIT2 = "<A CLASS=CC HREF='BSQORGANIGRAMA.ASP?BSQ=PORPUESTO&nid=" & nID & "' >BUSCAR&nbsp;POR&nbsp;PUESTO</A>"
             BTIT3 = "<A CLASS=CC HREF='BSQORGANIGRAMA.ASP?BSQ=PORNOMBRE&nid=" & nID & "' >BUSCAR&nbsp;POR&nbsp;NOMBRE</A>"
             BTIT4 = "<A CLASS=CC HREF='#' onclick=""mostrar('ayuda');"" >&nbsp;&nbsp;&nbsp;LEYENDA</A>"
             SACCION="corprint.asp?ID=XX"
			 BTIT1e = "<input  type=BUTTON  size=10%  value='IMPRIMIR' class=bietx  xSTYLE='cursor:pointer;color:"&session("CMFMEFT")&";background-color:"& Session("CMFMENU") &"' onclick='openw();'   >"
             'BTIT1e = BTIT1e & OFV.genform("TestFrm12", "AAA STYLE='display:none;'  target=_NEW", sAccion, "", "POST", "", "<input type=text NAME=PDF ID=PDF >", "si")

             SACCION="BSQORGANIGRAMA.ASP?BSQ=PORDEPENDENCIA&nid=" & nID
             BTIT1 = "<input  type=BUTTON  size=10%  value='BUSCAR&nbsp;POR&nbsp;DEPENDENCIA' class=bietx  xSTYLE='cursor:pointer;color:"&session("CMFMEFT")&";background-color:"& Session("CMFMENU") &"' onclick='document.all.TestFrm1.submit();'   >"
             BTIT1 = BTIT1 & OFV.genform("TestFrm1", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")

   	     SACCION="BSQORGANIGRAMA.ASP?BSQ=PORPUESTO&nid=" & nID
             BTIT2 = "<input  type=BUTTON  size=10%  value='BUSCAR&nbsp;POR&nbsp;PUESTO' class=bietx  xSTYLE='cursor:pointer;color:"&session("CMFMEFT")&";background-color:"& Session("CMFMENU") &"' onclick='document.all.TestFrm2.submit();'   >"
             BTIT2 = BTIT2 & OFV.genform("TestFrm2", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")


             SACCION="BSQORGANIGRAMA.ASP?BSQ=PORNOMBRE&nid=" & nID
	      BTIT3 = "<input  type=BUTTON  size=10%  value='BUSCAR&nbsp;POR&nbsp;NOMBRE' class=bietx  xSTYLE='cursor:pointer;color:"&session("CMFMEFT")&";background-color:"& Session("CMFMENU") &"' onclick='document.all.TestFrm3.submit();'   >"
             BTIT3 = BTIT3 & OFV.genform("TestFrm3", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")


            ' SACCION="BSQORGANIGRAMA.ASP?BSQ=PORDEPENDENCIA&nid=" & nID
           if session("APLMDCLI")="BC" THEN
	        BTIT4 = "<input  type=BUTTON  size=10%  value='LEYENDA' class=bietx  SxTYLE='cursor:pointer;color:"&session("CMFMEFT")&";background-color:"& Session("CMFMENU") &"' onclick=""mostrar('ayuda');""    >"
	       ELSE
	        BTIT4=""
	       END IF


           '  BTIT1 = BTIT1 & OFV.genform("TestFrm2", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")


             SEPARADOR="&nbsp;"
             LINEA2= "<v:polyline points=" & chr(34)
	     LINEA2=LINEA2 & " 0  0  40 0 " & chr(34)
	     LINEA2=LINEA2 & "   strokecolor=red ><v:fill on=""false"" /> <v:stroke dashstyle=""Dot"" weight=""1pt"" color=""BLACK"" /></v:polyline>"
         LINEA2="<p style='border-width:1px;border-style:dotted;'></p>"
	     LINEA3= "&nbsp;<b>Relación jer&aacute;rquica de Staff</b>"

         linea4="<p style='background:#F7EEDD; border-radius:25px;width:40pt;height:12pt'></p>"
	    ' linea4="<v:roundrect arcsize=""50%"" style=""width:30pt;height:15pt""><v:fill type=""solid"" color=""#F7EEDD"" /></v:roundrect>"
	     LINEA5="<b>Unidad de organización cuya jefatura está fuera de la misma</b>"
	     sxtabla="<table  align=left > <tr><td>"	 & linea2 & " </td><td>" & linea3 &"</td></tr>"
	     sxtabla=sxtabla & "<tr><td>"	& linea4	 & " </td><td>" & linea5 &"</td></tr></table>"

             ZTABLA=  "<br><TABLE      WIDTH=100% ><TR><TD  width=30% ><div id=ayuda style='display:none;'>"& sxtabla & " </td><td height=40px  align=right>"
             ZTABLA=ZTABLA  & BTIT4& SEPARADOR & BTIT1e  & SEPARADOR & BTIT1 &  SEPARADOR & BTIT2 &  SEPARADOR & BTIT3 &  SEPARADOR
             ZTABLA=ZTABLA & "</TD></TR></TABLE>"
             shtml=shtml & ZTABLA
            shtml = shtml & " <center><div id='myTreeContainer' style='width:100%;position:relative;top:10px;left:1px;' > "
            shtml = shtml & " </div> </center>"
            shtml = shtml & " <form id=ejecutar style='display:none' method=post></form>"


            shtml = shtml & "<script>"
            shtml = shtml & " var myTree = null;"
            delim = ""
            Amaxlevel = ""
            '  for i=0 to 8
            '  select case AposLevel(i)
            '  case  "A"
            '   Amaxlevel=Amaxlevel & DELIM & "10"
            '  case  "B"
            '   Amaxlevel=Amaxlevel & DELIM & "110"
            '  case  "C"
            '   Amaxlevel=Amaxlevel & DELIM & "210"
            '  case  "D"
            '   Amaxlevel=Amaxlevel & DELIM & "310"
            '  END SELECT
            '  DELIM=","
            '  next

            ' RESPONSE.WRITE Amaxlevel
            shtml=shtml &  "var aasis=new Array("&idsgen&");"


            shtml = shtml & " function CreateTree() { "
            shtml = shtml & "myTree = new ECOTree('myTree','myTreeContainer');"
            '''    shtml=shtml &  "myTree.config.colorStyle = ECOTree.CS_LEVEL;"
            shtml = shtml & "myTree.config.colorStyle = ECOTree.CS_NODE;"
            shtml = shtml & "myTree.config.nodeFill = ECOTree.NF_FLAT;"
            shtml = shtml & "myTree.config.useTarget = true;"
            shtml = shtml & "myTree.config.selectMode = ECOTree.SL_NONE;"
            shtml = shtml & "myTree.config.defaultNodeWidth = 85;"
            shtml = shtml & "myTree.config.defaultNodeHeight = 60;"
            shtml = shtml & "myTree.config.iSubtreeSeparation = 40;"
            shtml = shtml & "myTree.config.iSiblingSeparation = 20;"
            shtml = shtml & "myTree.config.iLevelSeparation = 50;"
           ' shtml = shtml & "myTree.config.nodeColor = " & Chr(34) & "#77AADD" & Chr(34) & ";"
           shtml = shtml & "myTree.config.topXAdjustment = 265;"
             shtml=shtml &  "myTree.config.rootXOffset =150; "
            shtml = shtml & "myTree.config.linkColor = 'black' ; "
            shtml = shtml & "myTree.config.iNodeJustification =ECOTree.NJ_CENTER; "

            shtml = shtml & wwwsettree

            shtml = shtml & "myTree.UpdateTree();"
            shtml = shtml & "}"

            shtml = shtml & " var elem = document.getElementById('myTreeContainer');"
		    shtml = shtml & "elemLeft = elem.offsetLeft;"
	        shtml = shtml & "elemTop = elem.offsetTop;"
            shtml = shtml & " CreateTree();elements = []; "

            shtml = shtml & " for (var k = 0; k < myTree.nDatabaseNodes.length; k++)"
		    shtml = shtml & "{  "
			shtml = shtml & "	node = myTree.nDatabaseNodes[k];"
			shtml = shtml & "	elements.push({width: node.w,"
    		shtml = shtml & "	height: node.h,"
    		shtml = shtml & "	top: node.YPosition,"
   			shtml = shtml & "	 left: node.XPosition,"
    		shtml = shtml & "	 url:node.target"
			shtml = shtml & "	});"
	        shtml = shtml & "var g = document.createElement('div'); g.style.width= node.w + 'px'; g.style.height=node.h + 'px'; g.style.backgroundColor='';g.style.position='absolute';"
			shtml = shtml &  " g.style.left=elemLeft + node.XPosition + 'px'; g.style.top=elemTop + node.YPosition + 'px';"
			shtml = shtml & "g.id= k + 'cmd';"
			shtml= shtml  &  "g.addEventListener('mouseover', mouseOver (g), false);"
			shtml= shtml  &  "g.addEventListener('mouseout', mouseOut(g), false);"
			shtml= shtml  &  "g.onclick=function ejecf(){ var xxx= this.id.replace(/[^0-9]/g,'');var fx=document.getElementById('ejecutar'); fx.action=elements[xxx].url;fx.submit(); ;return false;};"
			 shtml = shtml & "document.body.appendChild(g);"


			shtml = shtml & "}"

              shtml=shtml & " function mouseOver (e) {e.style.cursor='hand';e.style.cursor='pointer'};"
		      shtml=shtml & " function mouseOut (e) {e.style.cursor='auto;'};"
              'shtml=shtml & " function ejecf (e) { alert(e);};"
'
	  '     shtml = shtml & " var elem = document.getElementById('myTreeContainer');"
	  '     shtml = shtml & "elemLeft = elem.offsetLeft;"
	  '     shtml = shtml & "elemTop = elem.offsetTop;"
'




	  '     shtml = shtml & "elem.addEventListener('click', function(event) {"
	  '    shtml = shtml & " var x = event.pageX - elemLeft,"
	  '     shtml = shtml & "    y = event.pageY - elemTop;"
'
	  '     shtml = shtml & " elements.forEach(function(element) {  "
	  '       shtml = shtml & "  if (y > element.top && y < element.top + element.height && x > element.left && x < element.left + element.width) {"
	  '        shtml = shtml & "      var fx=document.getElementById('ejecutar'); fx.action=element.url;fx.submit();"
	  '      shtml = shtml & "   }"
	  '    shtml = shtml & " });"
'
      '   shtml = shtml & "}, false);"

	  shtml = shtml & "    function mostrar(showHideDiv) { "
	  shtml = shtml & "    	var ele = document.getElementById(showHideDiv);"

  shtml = shtml & "	    	if(ele.style.display == 'none') {"
  shtml = shtml & "	            ele.style.display = 'block';"

  shtml = shtml & "	      	}"
  shtml = shtml & "	    	else {"
  shtml = shtml & "	    		ele.style.display = 'none';"

  shtml = shtml & "	    	}"
  shtml = shtml & "	    } "



            shtml = shtml & " </script>"
           '' shtml = shtml & "<P>AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</P>"
            XXXX = GETOBJALOGUPD("ACCMOD")


        Else


        End If

        Call ofv.cerrarconn(cn)
        Call ofv.cerrarconn(cn2)
        Call ofv.cerrarconn(cn3)


        If request.querystring("mseg") = "M" Then
            shtml = shtml & "</body></html>"
        Else
        shtml = shtml & "<br> "
        shtml = shtml & "</body></html>"
        End If


        Generar = shtml

    End If





End If

'for i=0 to 8
'response.write  AposLevel(i)  & "   " &i & "<br>"
'next

Function buscapadre(idp)
    SPUESTO = idp
    SPUESTOx = ""
    StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
    StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid and listadepuestos.site='" & SESSION("SUN") &"') "
    StrSql = StrSql & " WHERE RelaPuesto.puestodepen = 0" & SPUESTO & "  ORDER BY RelaPuesto.puestorel "
    '  RESPONSE.WRITE STRSQL
    Set rsX = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not rsX.EOF Then

        Do While Not rsX.EOF
            SPUESTOx = rsX("PUESTOMASTER")
            If rsX("continuador") = "NO" Then Exit Do
            Call ofv.cerrarconsulta(rsX)
            StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
            StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid and listadepuestos.site='" & SESSION("SUN") &"') "
            StrSql = StrSql & " WHERE RelaPuesto.puestodepen = 0" & SPUESTOx & "  ORDER BY RelaPuesto.puestorel "

            Set rsX = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        Loop
    End If
    If SPUESTOx = "" Then SPUESTOx = SPUESTO
    buscapadre = SPUESTOx
End Function


Function AGREGARNIVEL(NIVEL, PADREx)


    For IX = 1 To NIVEL
        xin = xin + 1

        wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & ","& PADREx &",'virtualD'," & wwwsettreeAN & "," & wwwsettreeAL & ",'OrganigramaPuestosgraf.asp?ID=17&Sector=GERENTE_DE_AREA_RIESGO_CREDITICIO','#000000','',null,'0','SI',"&xin&");"

        PADREx = xin + 1
        xin = xin + 1
    Next

    AGREGARNIVEL = PADREx
End Function

Response.Write shtml


%>
