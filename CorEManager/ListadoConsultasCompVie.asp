
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = OFV.CheckUsuario()
If ok <> "Y" Then
    uv = OFV.uv
    If uv = "NO" Then ok = OFV.ValidarUsuario
    sHTML = ok
Else
    response.expires = 0

    Server.scripttimeout = 600
    response.buffer = True
    sMenu = Request.querystring("Menu")


    CantRegMostrar = 12
    CantRegMover = CDbl(Request.querystring("CantRegMover"))
    OrdenarPor = Request.querystring("OrdenarPor")
    sNombre = Request.querystring("Nombre")

    snombrev = Request.querystring("Nombre")
    smenuv = Request.querystring("Menu")

    If Len(Trim(OrdenarPor)) = 0 Then
        OrdenarPor = "DESCRI"
    End If

    OrdenarPor = UCase(OrdenarPor) & ""
    sOrdenx = Trim(OrdenarPor)
    ' response.write OrdenarPor
    Select Case OrdenarPor
        Case ""
            OrdenarPor = " order by complementos.descri"

        Case "DESCRI"
            OrdenarPor = " order by complementos.descri"
        Case Else
            OrdenarPor = " order by " & OrdenarPor
    End Select



    sMenu = UCase(sMenu)

    If sMenu = "PORTEXTO" Then

        If Request.Form.Count > 0 Then
            sIDv = Request.Form("buscar")

            sID = Replace(Request.Form("buscar"), "'", "")
            sID = Replace(sID, Chr(34), "")
        Else
            sIDv = Request.querystring("buscar")
            sID = Request.querystring("buscar")
            sID = OFV.convertircar(sID, "_", " ")
        End If

    ElseIf sMenu = "PORIDENTIFICACION" Then

        If Request.Form.Count > 0 Then

            sID = Request.Form("buscar")
            sIDv = Request.Form("buscar")
        Else
            sID = Request.querystring("buscar")
            sIDv = Request.querystring("buscar")
        End If

    Else

        If Request.Form.Count > 0 Then

            sID = Request.Form("lista")
            sIDv = Request.Form("lista")
        Else
            sID = Request.querystring("buscar")
            sIDv = Request.querystring("buscar")
        End If

    End If

    If sID = "" Then
        sIDv = Request.querystring("buscar")
        sID = Request.querystring("buscar")
        sID = OFV.convertircar(sID, "_", " ")
    End If


    sID = OFV.convertircar(sID, "_", " ")
    Set cn = OFV.conectar(OFV.strconn2)


    'Response.write smenu & "<br>"
    Select Case UCase(sMenu)
        Case "PORTIPO"

            StrSql = "Select DISTINCT Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND tipocomp='" & sID & "' " & OrdenarPor
            '    Response.write strsql & "<br>"
            sTitulo = "Componente: " & sID

        Case "PORTEXTO"
            sTitulo = "Texto: " & sID

            StrSql = "Select DISTINCT Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND Complementos.DESCRI  like '%" & sID & "%' " & OrdenarPor
            '  Response.write strsql & "<br>"
            sTitulo = "Texto: " & sID
        Case "PORIDENTIFICACION"

            StrSql = "Select DISTINCT Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND codiden ='" & sID & "'  " & OrdenarPor
            '   Response.write strsql & "<br>"
            sTitulo = "IDENTIFICACION " & sID

        Case "PORSISTEMA"
            StrSql = "select DISTINCT Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 from Compprty,complementos,TIPOCOMP"
            StrSql = StrSql & " where sistorigen='" & sID & "' "
            StrSql = StrSql & " and Compprty.codcomp=complementos.codcomp "
            StrSql = StrSql & " and Compprty.tipocomp=complementos.tipocomp "
            StrSql = StrSql & " and complementos.tipocomp=TIPOCOMP.TIPO  " & OrdenarPor
            ' Response.write strsql & "<br>"
    End Select

    ' Response.write strsql & "<br>"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")

    sHTML = OFV.FormHeader(sTitulo)
    xs = "<script language=JavaScript>" & Chr(13)

    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)

    xs = xs & "function Axn(c1,c2,c3) {" & Chr(13)

    xs = xs & " var mydat= 'EditarImagen2.asp?ID=' + c1 ; "

    xs = xs & " myWin= open(mydat"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)

    xs = xs & "</script>" & Chr(13)

    sHTML = sHTML & xs

    response.write sHTML
    sHTML = ""
    response.flush


    If session("WEBDESK") = "S" Then

    Else


        sID = OFV.convertircar(sID, " ", "_")
        Celda = OFV.GenCelda("", "ttr align=center ", "", "", "", "", "", "Tipo", "si")
        sLink = "<a class='nt' href=ListadoConsultasComp.asp?menu=" & sMenu & "&OrdenarPor=Descri&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Descripción</a>"
        Celda = Celda & OFV.GenCelda("", "ttr  align=center ", "", "", "", "", "", sLink, "si")
        sLink = "<a class='nt' href=ListadoConsultasComp.asp?menu=" & sMenu & "&OrdenarPor=descri2&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Tipo&nbsp;Componente</a>"
        Celda = Celda & OFV.GenCelda("", "ttr  align=center ", "", "", "", "", "", sLink, "si")
        sLink = "<a class='nt' href=ListadoConsultasComp.asp?menu=" & sMenu & "&OrdenarPor=codiden&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Identificación</a>"
        Celda = Celda & OFV.GenCelda("", "ttr  align=center  ", "", "", "", "", "", sLink, "si")
        Fila = OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

        response.write sHTML
        response.flush

        sHTML = ""


    End If

    nLinea = 0
    PosReg = 0
    On Error Resume Next

    dFechai = CLng(Date) - CLng(CDate("01-01-1900"))
    xsclase = " CLASS=VTI STYLE='background-color:#ffaa88;color:#000000;font-size:8pt;' "


    xTrColor = "#FFFFFF"


    If Not rs.EOF Then
        If session("WEBDESK") = "S" Then

            Do While (Not rs.EOF)
                XTABLAN = ""

                '''''''''''''''''''''''''''''''''''''''''''''''''''
                XTABLAN = XTABLAN & Chr(13)
                XTABLAN = XTABLAN & "<div id='DTNOVINT" & rs(0) & "K' style='float: left; display: inline;  "
                XTABLAN = XTABLAN & "border-style:solid;border-width:4;border-color:#ffffff;' >"


                ''   tabprocesos (iproc,iprdoc,1) = sitem                        'nombre
                ''   tabprocesos (iproc,iprdoc,2) = rs(0)                        'codigo
                ''   tabprocesos (iproc,iprdoc,3) = sPerfil                      'perfil
                ''   tabprocesos (iproc,iprdoc,4) = "Documento normativo"        'copete
                ''   tabprocesos (iproc,iprdoc,5) = rs(6)                        'fec.vigencia
                ''   tabprocesos (iproc,iprdoc,6) = 0                            'corporativo 0-no 1-si
                ''   tabprocesos (iproc,iprdoc,7) = stitem                       'stitem
                ''   tabprocesos (iproc,iprdoc,9) = rs(19)                        'objimagen
                ''   tabprocesos (iproc,iprdoc,10) = rs(20)                        'objurldsc
                ''   tabprocesos (iproc,iprdoc,11) = rs(21)                        'tema
                ''   tabprocesos (iproc,iprdoc,12) = rs(22)                        'subtema



                wspwkcopete = ""
                wspwkimagen = ""
                wspwkurldsc = ""

                StrSql = "SELECT g.objcopete, g.objimagen, g.objurldsc   "
                StrSql = StrSql & " from objinfo g where g.codobjeto = 0" & rs(0)
                StrSql = StrSql & " and g.tipobjeto = '" & rs(2) & "' "
                Set rsxx1 = OFV.crearconsultaEx(StrSql, cn, 1, "")
                If Not rsxx1.EOF Then
                    wspwkcopete = rsxx1(0) & ""
                    wspwkimagen = rsxx1(1) & ""
                    wspwkurldsc = rsxx1(2) & ""
                End If


                wtabprocesos001 = rs(3)                                                         'nombre
                wtabprocesos002 = rs(0)                                                         'codigo
                wtabprocesos003 = "perfil"                                                      'perfil
                wtabprocesos004 = "Componente "                                                 'copete
                wtabprocesos005 = CLng(Date) - CLng(CDate("01-01-1900"))                                                  'fec.vigencia

                wtabprocesos006 = 0                                                             'corporativo 0-no 1-si
                wtabprocesos007 = rs(2)                                                         'stitem
                wtabprocesos008 = CLng(Date) - CLng(CDate("01-01-1900"))                                                 'fec.vigencia
                wtabprocesos009 = wspwkimagen                                                   'objimagen
                wtabprocesos010 = wspwkurldsc                                                   'objurldsc
                wtabprocesos011 = rs(1)                                                         'tema
                wtabprocesos012 = rs(6)                                                         'subtema


                If wspwkcopete = "" Then
                    wtabprocesos004 = "Componente "                                                      'copete
                Else
                    wtabprocesos004 = Capitalize(LTrim(wspwkcopete))                                     'copete
                End If


                xtmsgIMG = wtabprocesos009 & ""
                xtmsgVID = wtabprocesos010 & ""
                xtmcod = wtabprocesos002 & ""

                xtmdtema = wtabprocesos011 & ""
                xtmdstema = wtabprocesos012 & ""

                xtmsgtit = Capitalize(wtabprocesos001)

                If (dFechai - wtabprocesos008) < 31 Then
                    '      xtmsgtit = "<B><font color='#ffffff' STYLE='background-color:#ff0000;' size=1 >&nbsp;nuevo&nbsp;</font></B>&nbsp;&nbsp;" & xtmsgtit
                End If

                xtmsgCOP = wtabprocesos004


                ''   If tabprocesos006 = 1 Then
                ''      smasc = "&nbsp;&nbsp;&nbsp;<font color=#993333 ><font class=btic title='Corporativo' >G</font></font>"
                ''   Else
                smasc = "&nbsp;"
                ''   End If


                Celda = ""

                sImage = "<font face='webdings' size=4 color='#005500' >&#0157;</font>"
                sHref = "&nbsp;"
                sFuncion = " onclick=" & Chr(34) & " Axn('" & wtabprocesos002 & "','" & wtabprocesos007 & "','" & wtabprocesos003 & "'); return false;" & Chr(34)

                sLink1 = "&nbsp;<font color='#AA0000' size=2 >&#0149;</font>&nbsp;"
                sLink = "&nbsp;&nbsp;" & wtabprocesos001


                '''''''''''''''
                XTABLAN = XTABLAN & "<table  cellspacing=0 cellpadding=1  style='width:350px;border-style:solid;border-width:1;border-color:#666666;' >"
                XTABLAN = XTABLAN & "<tr HEIGHT=25  >"
                XTABLAN = XTABLAN & "<td " & xsclase & " style='cursor:hand;' title='Consultar.' " & sFuncion & "  >" & xtmsgtit & "</TD>"
                XTABLAN = XTABLAN & "</TR>"
                XTABLAN = XTABLAN & "</table>"

                ''''''''''''''''''
                XTABLAN = XTABLAN & "<table cellspacing=0 cellpadding=0 style='width:350px;border-style:dotted;border-width:1;border-color:#bbbbbb;"
                XTABLAN = XTABLAN & "border-top-style:solid;border-top-width:1;border-top-color:#ffffff;' >"

                XTMSGSEND = ""
                XTMSGSEND = XTMSGSEND & "<font color=#555555 style='font-weight:bold;' size=1 >" & xtmdstema & "</font>"

                XTABLAN = XTABLAN & "<tr >"
                XTABLAN = XTABLAN & "<td class=UT align='right' colspan=8 STYLE='background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
                XTABLAN = XTABLAN & "</TR>"

                XTABLAN = XTABLAN & "<tr >"
                XTABLAN = XTABLAN & "<td class=UT width='5%'  >&nbsp;</TD>"
                For yy = 1 To 6
                    XTABLAN = XTABLAN & "<td class=UT width='15%'  >&nbsp;</TD>"
                Next
                XTABLAN = XTABLAN & "<td class=UT width='5%'  >&nbsp;</TD>"
                XTABLAN = XTABLAN & "</TR>"


                If (xtmsgIMG & "") = "1" Then


                    StrSql = "Select * From objinfoIMG  where codobjeto = 0" & xtmcod & " ORDER BY objSECUEN, OBJID"
                    Set RS1 = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not RS1.EOF Then

                        XTABLAN = XTABLAN & "<tr bgcolor='#F0F0F0' >"
                        XTABLAN = XTABLAN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

                        XTABLAN = XTABLAN & "<table cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#990000;' >"
                        XTABLAN = XTABLAN & "<tr   >"

                        XSFSTYLE = " STYLE='display:block;' "

                        Do While Not RS1.EOF
                            XTABLAN = XTABLAN & "<td CLASS=UT " & XSFSTYLE & " ><IMG SRC='" & RS1(6) & "' ></TD>"
                            XSFSTYLE = " STYLE='display:none;' "

                            RS1.movenext
                        Loop

                        XTABLAN = XTABLAN & "</TR>"
                        XTABLAN = XTABLAN & "</table>"

                        XTABLAN = XTABLAN & "</TD>"
                        XTABLAN = XTABLAN & "</TR>"

                        XTABLAN = XTABLAN & "<tr HEIGHT='20' >"
                        XTABLAN = XTABLAN & "<td class=UT colspan=8  >&nbsp;</TD>"
                        XTABLAN = XTABLAN & "</TR>"

                    End If

                ElseIf (xtmsgIMG & "") = "2" Then

                    XTABLAN = XTABLAN & "<tr bgcolor='#F0F0F0' >"
                    XTABLAN = XTABLAN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
                    XTABLAN = XTABLAN & xtmsgVID
                    XTABLAN = XTABLAN & "</TD>"
                    XTABLAN = XTABLAN & "</TR>"

                    XTABLAN = XTABLAN & "<tr HEIGHT='20' >"
                    XTABLAN = XTABLAN & "<td class=UT colspan=8  >&nbsp;</TD>"
                    XTABLAN = XTABLAN & "</TR>"

                End If

                XTMSGSEND = "<B><font color=#888888 >" & gtfechadorsh(wtabprocesos005) & "</font></B>"

                XTMSGSEND = XTMSGSEND & "<B><font color='#aa0000' size=2 >&#0124;</font></B>"
                XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-weight:bold;' >&nbsp;" & xtmdtema & "&nbsp;</font>"

                XTABLAN = XTABLAN & "<tr height=25 >"
                XTABLAN = XTABLAN & "<td class=uti colspan=7 STYLE='color:#000000;' >"
                XTABLAN = XTABLAN & XTMSGSEND & "</TD>"
                XTABLAN = XTABLAN & "<td class=ut  >&nbsp;</TD>"
                XTABLAN = XTABLAN & "</TR>"

                XTMSGSEND = ""
                XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-size:8pt;font-weight:normal;' >" & xtmsgCOP & "</font>"

                XTABLAN = XTABLAN & "<tr  >"
                XTABLAN = XTABLAN & "<td class=ut VALIGN='TOP' align='center' ><font color='#AA0000' size=2 >&#0149;</font></TD>"
                XTABLAN = XTABLAN & "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
                XTABLAN = XTABLAN & XTMSGSEND & "</TD>"
                XTABLAN = XTABLAN & "<td class=ut  >&nbsp;</TD>"
                XTABLAN = XTABLAN & "</TR>"

                XTABLAN = XTABLAN & "<tr >"
                XTABLAN = XTABLAN & "<td class=UT colspan=8  >&nbsp;</TD>"
                XTABLAN = XTABLAN & "</TR>"

                XTABLAN = XTABLAN & "</table>"
                XTABLAN = XTABLAN & "</div >"

                XTABLAN = XTABLAN & Chr(13)
                sHTML = sHTML & XTABLAN

                rs.movenext

            Loop
        Else

            xTrColor = "#FFFFFF"

            If Not rs.EOF Then rs.Move (CantRegMover)

            Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

                sImagen = "<font face=wingdings size=4 color=#993333>2</font>"

                sHref2 = "EditarImagen2.asp?ID=" & rs(0)
                sHref = "javascript:void(0);"


                sLink = OFV.GenLink("", "nt onclick=""AbrirCons('" & sHref2 & "');return false;"" ", sHref, "_blank", "", "", sImagen, "si")
                Celda = OFV.GenCelda("", "ttr", "center", "2%", "", "", "", sLink, "si")


                sLink = OFV.GenLink("", "nt onclick=""AbrirCons('" & sHref2 & "');return false;"" ", sHref, "_blank", "", "", "<u>" & rs(3) & "</u>", "si")
                Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", sLink, "si")
                Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", rs(6), "si")
                Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", rs(1), "si")

                Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

                If xTrColor = "#FFFFFF" Then
                    xTrColor = "#F9F9F9"
                Else
                    xTrColor = "#FFFFFF"
                End If




                PosReg = PosReg + 1
                rs.movenext
                nLinea = nLinea + 1

                If nLinea > 20 Then
                    nLinea = 0
                    response.flush
                End If



            Loop


        End If


        sAccion = "SubconsultaComponentes.asp?Menu=" & sMenu & "&Nombre=" & sNombre
        Celda = OFV.BotonVolver(sAccion)

        If session("WEBDESK") = "S" Then

            sHTML = sHTML & OFV.botoneraESSX("0", "", "", "", "no", "no", "", Addqry, "", "no", "", "", Celda)

        Else

            sHTML = sHTML & OFV.GenTabla("", "", "", "90% ", "", "0", "", "1", "2", "", "", Fila, "si")

            Addqry = "&Menu=" & smenuv & "&OrdenarPor=" & sOrdenx
            Addqry = Addqry & "&buscar=" & sIDv & "&nombre=" & snombrev

            sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)

        End If

        sHTML = sHTML & "</body></html>"

        Call OFV.cerrarconsulta(rs)
        Call OFV.cerrarconn(cn)

        response.write sHTML
        response.flush

        sHTML = ""
    Else
	
        sAccion = "SubconsultaComponentes.asp?Menu=" & sMenu & "&Nombre=" & sNombre
        Celda = OFV.BotonVolver(sAccion)

        Addqry = "&Menu=" & smenuv & "&OrdenarPor=" & sOrdenx
        Addqry = Addqry & "&buscar=" & sIDv & "&nombre=" & snombrev

        sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)

		Celda   = ""

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        'sLink   = "No se encuentran documentos para esta busqueda."
		sLink = "<BR><BR>&nbsp;<B><font color=#993333 SIZE=3 FACE='WEBDINGS' >i</font></B>"
        sLink = sLink &  "&nbsp;NO SE ENCONTRARON DOCUMENTOS PARA LA BUSQUEDA SOLICITADA<BR><BR>"
        Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","",sLink,"si")

        Fila    = Fila & ofv.GenRow("","aaa bgcolor='" & xTrColor & "' ","","","25","","",Celda,"si")
		
        sHTML = sHTML & Fila
		
        sHTML = sHTML & "</body></html>"

        Call OFV.cerrarconsulta(rs)
        Call OFV.cerrarconn(cn)

        response.write sHTML
        response.flush

        sHTML = ""
    End If
End If

response.write sHTML


%>
