<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok = ofv.CheckUsuario()
If ok <> "Y" Then
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    sHTML = ok
Else

    response.expires = 0
    Server.ScriptTimeout = 120

    SMODAL = request.querystring("modal")
    sTipo = request.querystring("Tipo")
    sNombre = request.querystring("Nombre")
    sNombre = ofv.convertircar(sNombre, "_", " ")
    sMenu = request.querystring("Menu")
    sMETODO = request.querystring("METODO")
    sACTIVIDAD = request.querystring("ACTIVIDAD")
    SACCIONEJEC = request.querystring("ACCIONEJEC")
    SACCIONEJECCOD = request.querystring("ACCIONEJECCOD")
    SACCIONEJECTIP = request.querystring("ACCIONEJECTIP")
    SVARIABLE = UCase(CStr(request.querystring("OPCIONES")))
	SPROCESO = REQUEST.QueryString("QSPROCESO")
	SPERIODO = REQUEST.QueryString("QSPERIODO")
	IF (ISNULL (SPERIODO) OR SPERIODO = "") THEN
		SPERIODO = "ALL"
	END IF

    If SMODAL = "" Then SMODAL = "LISTA"

    session("ModMain") = "ED"

    xapli = 1

    On Error Resume Next
    Set cn = ofv.conectar(ofv.strconn1)
    StrSql = "Select ID From CorAplic where codapli = '" & Trim(session("CorApli")) & "' "
    Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not rs.EOF Then
        xapli = rs(0)
    End If

    StrSql = "Select menu From menux where menuatr = '" & sNombre & "' and aplicacion = " & xapli
    Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not rs.EOF Then
        sNombre = rs(0)
    End If
    Call ofv.cerrarconsulta(rs)
    Call ofv.cerrarconn(cn)

    ''''''''''''''''''''''''''''

    session("xorigen") = "tema"

    XXXX = NEWEMPRE("NO")

    SEMPRE = request.querystring("EMPRE")

    If SEMPRE = "" Then
        SEMPRE = ATREMP(0, 0)
    Else
        recarga = NEWEMPRE(SEMPRE)
    End If

    sUN = SEMPRE

    If sUN = "" Then sUN = 1
    session("UN") = sUN

    LNombre = ofv.convertircar(sNombre, " ", "_")

    If sMenu <> "" Then session("Menu") = sMenu

    sID = request.querystring("ID")
    If sNombre = "" Then sNombre = UCase(sTipo)
    sitem = ofv.convertircar(request.querystring("sitem"), "_", " ")
    stitem = ofv.convertircar(request.querystring("stitem"), "_", " ")
    sticono = request.querystring("sticono")
    stnivel = request.querystring("stnivel")

    Set cn = ofv.conectar(ofv.strconn2)
    StrSql = "Select estado from estados where objid = " & session("sEstadon")
    ' response.write strsql & "<br>"
    Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not rs.EOF Then
        sstipo = rs(0)
    End If
    Call ofv.cerrarconsulta(rs)

    sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
    sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
    sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
    sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
    sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
    sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

    funcion = "  onload=" & Chr(34)

    If recarga = "S" And sTipo <> "Consulta" Then
        funcion = funcion & "document.recarga.submit();"
    End If

    funcion = funcion & "rszdx4();" & Chr(34)
    If SMODAL <> "LISTA" Then
        funcion = funcion & " onmousedown = " & Chr(34) & "Abre();" & Chr(34) & " "
    End If
    funcion = funcion & " onresize = " & Chr(34) & "rszdx4();" & Chr(34) & " "

    sHTML = ofv.MENUheader("", funcion)
    sHTML = sHTML & "<script language=javascript>" & Chr(13)
    sHTML = sHTML & "function mostrarmenuD(){" & Chr(13)
    sHTML = sHTML & "document.all.Dx1.style.display='block';" & Chr(13)
    sHTML = sHTML & "}" & Chr(13)
    sHTML = sHTML & "function rszdx4(){" & Chr(13)
    'shtml=shtml & "document.all.DX2.style.display='block';" & chr(13)
    sHTML = sHTML & "}" & Chr(13)
    sHTML = sHTML & "</script>" & Chr(13)

    If SMODAL <> "LISTA" Then
        sHTML = sHTML & "<script type=" & Chr(34) & "text/javascript" & Chr(34) & " src='IncFile/dragdrop2.js' ></script>"
    End If

    sHTML = sHTML & ofv.CorFormHeader

    sNombre = ofv.convertircar(sNombre, " ", "_")

    HrfAsp = "EMDETAPROCPROPER.asp?"
    HrfTrg = "fesquemas"

    Fila = ""
    Celda = ""
    sHref = ""

    If session("EMPMUN") = "S" Then

        Fila = ""

        SINPUT3 = "&nbsp;<A class=tt title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

        Fila = ""
        Celda = ""
        sHref = ""

        sImage = ofv.GenImage("", "", "", "", "..\CORESS20\EMPRESA\" & ATREMP(0, 5), "0", "")
        sLink = ofv.GenLink("", "ttt", sHref, "ftemas", "", "", sImage, "si")
        Celda = Celda & ofv.GenCelda("", "tt ", "", "8%", "", "", "", SINPUT3, "si")
        Celda = Celda & ofv.GenCelda("", "ttt ", "center", "5%", "", "", "", sLink, "si")
        Celda = Celda & ofv.GenCelda("", "ttt  ", "", "", "", "", "", "&nbsp;<BIG>" & ATREMP(0, 4) & "</BIG>", "si")

        Fila = Fila & ofv.GenRow("", "", "top", "", "", "", "", Celda, "si")

        TABLA = ofv.GenTabla("", "aaa style='" & session("CBRDABL") & "' ", "", "100% ", "", "0", "", "-1", "-2", "", "", Fila, "si")

        sHTML = sHTML & "<script language=javascript>" & Chr(13)
        sHTML = sHTML & "function shwDem(){" & Chr(13)
        sHTML = sHTML & "document.all.DEM1.style.display='block';" & Chr(13)
        sHTML = sHTML & "}" & Chr(13)
        sHTML = sHTML & "</script>" & Chr(13)
        '----------------------------------------------
        sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
        sinput2 = sinput2 & "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
        sinput2 = sinput2 & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "document.all.DEM1.style.display='none';" & Chr(34)
        sinput2 = sinput2 & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
        TABLA = TABLA & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:150px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"

        TablaX = TablaX & ofv.GenTabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
        TablaX = TablaX & "<tr><td  "
        TablaX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
        TablaX = TablaX & "<tr><td colspan=3 "
        TablaX = TablaX & " >&nbsp;</td></tr>"

        For IAEMP = 1 To TAEMP
            '''  RESPONSE.WRITE  ATREMP(IAEMP,2)
            If ATREMP(IAEMP, 2) = 1 Then

                If ATREMP(IAEMP, 0) = ATREMP(0, 0) Then
                    TablaX = TablaX & "<tr BGCOLOR=#EEEEEE >"
                Else
                    TablaX = TablaX & "<tr>"
                End If

                TablaX = TablaX & "<td width=5 "
                If ATREMP(IAEMP, 0) = ATREMP(0, 0) Then
                    TablaX = TablaX & "class=vt  >" & sImagen4si & "</td>"
                Else
                    TablaX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
                End If

                sImage = ofv.GenImage("", "", "", "", "..\CORESS20\EMPRESA\" & ATREMP(IAEMP, 5), "0", "")
                TablaX = TablaX & "<td width=5 "
                TablaX = TablaX & "class=vt  >" & sImage & "</td>"

                TablaX = TablaX & "<td class=vt  >"

                sHref = "OBJtree.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & session("MENU") & "&EMPRE=" & ATREMP(IAEMP, 0)
                sLink = ofv.GenLink("", "cc", sHref, "", "", "", "<b>" & ATREMP(IAEMP, 4) & "</b>", "si")
                TablaX = TablaX & sLink
                TablaX = TablaX & "</td>"
                TablaX = TablaX & "</tr>"
            End If

        Next

        TablaX = TablaX & ofv.genFtabla()

        TABLA = TABLA & TablaX
        TABLA = TABLA & "</DIV>"
        sHTML = sHTML & TABLA
    End If

    Fila = ""
    TABLA = ""

    sHref = ""
    sImage = ofv.GenImage("", "", "", "", "Iconos\ICorMMenu.jpg", "0", "")

    sImage = ofv.GenImage("", "", "", "", "..\CORESS20\EMPRESA\" & ATREMP(0, 5), "0", "")

    sLink = ofv.GenLink("", "nt", sHref, "ftemas", "", "", sImage, "si")
    Celda = ofv.GenCelda("", "aaa colspan=3 rowspan=2  style='" & session("CBRDICO") & "'", "center", "2%", "", "", "", sLink, "si")

    xxx = "<b><u>" & UCase(sstipo) & "</u></b>"

    sLink = ofv.GenLink("", "cc", sHref, "ftemas", "", "", "<b><u>" & ATREMP(0, 4) & "</u></b>", "si")
    Celda = Celda & ofv.GenCelda("", "aaa rowspan=2", "", "", "", "", "", "&nbsp;", "si")
    Celda = Celda & ofv.GenCelda("", "aaa colspan=3 rowspan=2", "", "", "", "", "", sLink, "si")
    Celda = Celda & ofv.GenCelda("", "", "", "", "", "", "", SINPUT3, "si")

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' set de permisos
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''''''''''''''''''''''''''''''''''''''''''
    If session("DELLOG") = "S" Then

        XMFTGRP = Split(session("SegGrpLst"), Chr(9))
        xcount = UBound(XMFTGRP) + 1
        XMFSEP = ""
        WKSPSDKGRP = ""

        For XMFI = 0 To UBound(XMFTGRP)
            WKSPSDKGRP = WKSPSDKGRP & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
            XMFSEP = ","
        Next

    Else
        WKSPSDKGRP = "'GRP-" & session("Seggrp") & "'"
    End If


    ''''''''''''''''''''''''''''''''''''''''''

    stitem = "Tema"
    sIcono = "ICorTema.jpg"
    sNuevoTipo = "SUBTEMAS"
    nNivel = 0

    sID = 0
    nv = 0
    cborderB = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("APLSHW") & ";"
    cborderl = "border-left-style:solid;border-left-width:1;border-left-color:" & session("APLSHW") & ";"
    borderB = ""
    borderl = ""

    '''''''''''''''''''''''''''''''''''''''''''
    ''' session("menu") = 3
	
    scodmaster = request.querystring("codmaster") & ""
    stipmaster = request.querystring("tipmaster") & ""
    sCODNIVEL = request.querystring("CODNIVEL")

    If scodmaster = "" Or stipmaster = "PROCESO" Then
        scodmaster = "000"
        stipmaster = "PROCESO"
        sCODNIVEL = 1

        session("OROBJ") = 0
    End If

    session("ASP") = "vuelveprocesos.asp"
    session("ASPSTR") = wkproccod
    session("ASPSTRD") = session("ASPSTR")

    ''''''''''''''''''''''''''''''''''''''''''''

    sImageCRT = "<font face='corflaticon' size=2 color='#0000ff' style='cursor:pointer;border-style:OUTSET;border-width:1;' title='Tiene circuito definido.' >&#061;</font>"

    session("ORPERIODOST") = XFORPERIODOST
    session("ORRSKST") = 1

    wspTIPO = "SP"
    wkproccodL = "&nbsp;"
    wkproccodL2 = "&nbsp;"
    wkproccodN = "&nbsp;"

        If CLng("0" & request.Form("PERIODO")) > 0 Then
            XTPPER = CLng("0" & request.Form("PERIODO"))
        Else
            XTPPER = "ALL"
        End If

        XTPROC = CLng("0" & request.Form("PROC"))

    If (SVARIABLE <> "MOSTRAR" OR CSTR(SPERIODO) = "ALL") Then
		TABLA = ""

        Celda = ofv.GenCelda("", "tt", "", "", "", "", "", "PERIODO", "si")

        StrSql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
        StrSql = StrSql & " ESTADO IN ('OK','CL')  "
        StrSql = StrSql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"
        sInput = ofv.GenerarCombo("Subconsulta0", "PERIODO", XTPPER, StrSql, cn, 0, 1, "")

        Celda = Celda & ofv.GenCelda("", "vti", "", "", "", "", "", sInput, "si")
        Fila = Fila & ofv.GenRow("", "", "", "", "30", "", "", Celda, "si")

        sAccion = "ORESTPROCESOS.asp"
        Formulario = Formulario & ofv.GenForm("Subconsulta0", "", sAccion, "", "post", "", Fila, "si")

        Fila = ""

        '--------------------

        Celda = ofv.GenCelda("", "tt", "", "", "", "", "", "PROCESO", "si")

        StrSql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
        StrSql = StrSql & " AND codobjmaster = '" & scodmaster & "' and tipobjmaster = '" & stipmaster & "'  "
        StrSql = StrSql & " ORDER BY OBJETOS.INDOBJ "

        sInput = ofv.GenerarCombo("Subconsulta1", "PROC", XTPROC, StrSql, cn, 0, 3, "")
        sInput = sInput & ofv.GenerarInput("PERIODO", XTPPER, "hidden", "Subconsulta1", "1", "dt", "")
        Celda = Celda & ofv.GenCelda("", "vti", "", "", "", "", "", sInput, "si")
        Fila = Fila & ofv.GenRow("", "", "", "", "30", "", "", Celda, "si")

        sAccion = "ORESTPROCESOS.asp"
        Formulario = Formulario & ofv.GenForm("Subconsulta1", "", sAccion, "", "post", "", Fila, "si")

        Fila = ""
        '-----------------------------

        If XTPROC = 0 Then XTPROC = "ALL"

        Celda = ""
        Celda = Celda & ofv.GenCelda("", "vti COLSPAN=2", "", "", "", "", "", "&nbsp;", "si")
        Fila = Fila & ofv.GenRow("", "", "", "", "30", "", "", Celda, "si")

        formx = Fila

        sAccion = "ORESTPROCESOS.asp?OPCIONES=MOSTRAR&QSPERIODO=" & XTPPER & "&QSPROCESO=" & XTPROC

        Celda = ofv.GenCelda("", "ut", "", "", "", "", "", "&nbsp;", "si")
        sInput = ofv.GenerarInput("", "Buscar", "submit", "", "10", "bt", "")
        sInput = sInput & ofv.GenerarInput("Tcons", sMenu, "hidden", "", "1", "dt", "")
        Celda = Celda & ofv.GenCelda("", "vti", "Center", "", "", "", "", sInput, "si")
        formx = formx & ofv.GenRow("", "", "", "", "40", "", "", Celda, "si")

        Formulario = Formulario & ofv.GenForm("", "", sAccion, "", "post", "", formx, "si")
		
		TABLA = "<center> <BR>"

        TABLA = TABLA & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","60% ","","0","","5","0","","",Formulario,"si")
		
		TABLA = TABLA & "</center>"

    Else
'******************************************************************************************************************
'Comienzo de configuración del reporte
'******************************************************************************************************************
		TABLA = ""
		
        Dim TABSPPERM(100, 3)
        IPERM = 0
        TPERM = 0

        Dim TABMAPA(10, 10)

        TTABMAPA = 0

        StrSql = ""
        StrSql = StrSql & " SELECT * FROM ORPERIODOS "
        StrSql = StrSql & " WHERE ESTADO = 'OK' AND APLIC = 'OR'  "
		'RESPONSE.Write(SPERIODO)
		IF (CSTR(SPERIODO) <> "ALL") THEN
			StrSql = StrSql & " AND OBJID = " & SPERIODO
		END IF
        StrSql = StrSql & " ORDER BY INICIO DESC, OBJID DESC "
		
        'response.write (Chr(13) & Chr(13))
	    'response.write (StrSql)
    	'response.write (Chr(13) & Chr(13))
	
	
		Set RSX5 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not RSX5.EOF Then
            Do While Not RSX5.EOF

                TTABMAPA = TTABMAPA + 1
                TABMAPA(TTABMAPA, 0) = RSX5(0)
                TABMAPA(TTABMAPA, 1) = RSX5(1)
                TABMAPA(TTABMAPA, 2) = RSX5(2)
                TABMAPA(TTABMAPA, 3) = 0
                TABMAPA(TTABMAPA, 4) = 0
                TABMAPA(TTABMAPA, 5) = 0


                RSX5.movenext
            Loop
        End If


        wkprocSPW = "&nbsp;"
        WKSPPERMISOS = "&nbsp;"
        WKSPTIPOPR = "NO"

        WKSPMASTCT = "S"
        WKSPMASTMO = "S"
        WKSPMASTCO = "S"
        WKSPMASTCE = "S"


        WKSPCIRCCT = "N"
        WKSPCIRCMO = "N"
        WKSPCIRCCO = "N"
        WKSPCIRCCE = "N"

        WKPROCDESC = "PRINCIPAL DEL NEGOCIO&nbsp;(NIVEL&nbsp;" & sCODNIVEL & ")"
        wkproccod = "&nbsp;"
        wkproccod2 = "&nbsp;"
        wspTIPO = "SP"

        sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

        ''WKSPPERMISOS

        sInput = "PROCESO"
        wspTIPOCLC = "#993333"

        sImagePRC = "<font face='corflaticon' size=2 color='" & wspTIPOCLC & "'  >&#060;</font>"

        WKPROCDESC = WKDIVSECUEN & sInput & ":&nbsp;" & WKPROCDESC


        Celda = ""

        Celda = Celda & ofv.GenCelda("", "Ut style='background-color:#f0f0f0;color:#990000;' ", "CENTER", "1%", "", "", "", sImagePRC, "si")
        Celda = Celda & ofv.GenCelda("", "ut style='background-color:#f0f0f0;color:#888888;' ", "", "85%", "", "", "", WKPROCDESC, "si")
        'Celda = Celda & ofv.GenCelda("","ut style='background-color:#f0f0f0;color:#888888;' ","","85%","","","","hOLA A TODOS ","si")

        Fila A = ofv.GenRow("", "", "", "", "", "", "", Celda, "si")
        TablaH = ofv.GenTabla("", "aaa style='background-color:#f0eedd;border-style:solid;border-width:1;border-color:#000000;' valign='middle' align='center' ", "", "100% ", "30", "0", "", "-1", "1", "", "", Fila, "si")
        ''''''''''''''''''''''''''''''''''''''''''''
        TablaVH = ""
        ''''''''''''''''''''''''''''''''''''''''''''
        Fila = ""
        wspcolor1 = "#990000"
        wspcolor2 = "#009900"
        sImageO1 = "&#0253;"
        sImageO2 = "&#0254;"

        StrSql = "Select A.* " & CHR(13)
		StrSql = StrSql & " From OBJETOS A " & CHR(13)
        StrSql = StrSql & " INNER JOIN relaobj B ON ( B.CODOJASOC = A.CODOBJETO )" & CHR(13)
        StrSql = StrSql & " where  " & CHR(13)
		IF (SPROCESO = 0) THEN
		
	        StrSql = StrSql & "  B.codobjmaster = '" & scodmaster & "' and B.tipobjmaster = '" & stipmaster & "'  " & CHR(13)
		
		ELSE
			StrSql = StrSql & "  A.CODOBJETO = " & SPROCESO & CHR(13)
		END IF
        StrSql = StrSql & "  AND (A.CODVINOBJ = 'SP' or A.CODVINOBJ = 'SPI') " & CHR(13)
        StrSql = StrSql & " ORDER BY A.TSEGMENTO, A.TAGRUPA, A.INDOBJ " & CHR(13)

        'response.write (Chr(13) & Chr(13))
	    'response.write (StrSql)
    	'response.write (Chr(13) & Chr(13))
		
        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
		
        If rs.EOF Then

            Celda = ""
            Celda = Celda & ofv.GenCelda("", "vtG ", "CENTER", "", "", "", "", "NO HAY PROCESOS, ACTIVOS O ACTIVIDADES ASOCIADAS", "si")
            Fila = Fila & ofv.GenRow("", "", "", "", "", "", "", Celda, "si")
            TablaN = ofv.GenTabla("", "aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ", "", "300 ", "", "0", "", "8", "8", "", "", Fila, "si")

            TABLA = "<CENTER><BR><BR>" & TABLA & "<BR><BR>" & TablaN & "</CENTER>"

        Else

            xTrColor = "#FFFFFF"

            wkcant = 1
            wkleft = 5
            wktop = 5

            DHOY = CLng(Date) - CLng(CDate("01-01-1900"))


            Celda = ""
            Celda = Celda & ofv.GenCelda("", "utI COLSPAN=2 STYLE='COLOR:#FFFFFF;' ", "", "", "", "", "", "PROCESO", "si")
            Celda = Celda & ofv.GenCelda("", "UtI STYLE='COLOR:#FFFFFF;' ", "", "", "", "", "", "FRECUENCIA", "si")
            Celda = Celda & ofv.GenCelda("", "UtI COLSPAN=2 STYLE='COLOR:#FFFFFF;' ", "", "", "", "", "", "ULTIMA EVALUACION", "si")

            Filan = Filan & ofv.GenRow("", "AAA bgcolor='#888888' ", "", "", "20", "", "", Celda, "si")



            Do Until rs.EOF

                wkcoddelete = "N"
                wktitdelete = " "

                ''''''''''''''''''''''''''''''

                XFNOVECRT = ""
                XFNOVEPROC = ""
                WKSPESTADO = 3
                wspcolor = "#993333"
                sImageP = "&#0167;"
                wspcolor2 = wspcolor

                sImageP = sImagePER & "<font face='corflaticon' size=2 color='" & wspcolor2 & "'>" & sImageP & "</font>"

                wkproccod = rs(3)

                wkproccodFC = CLng("0" & rs(10))

                Select Case (rs(9) & "")
                    Case "1"
                        wkproccodF = "<B>" & wkproccodFC & "&nbsp;MES/ES" & "</B>"
                        WSPCANTDIAS = wkproccodFC * 30
                    Case "2"
                        wkproccodF = "<B>" & wkproccodFC & "&nbsp;AÑO/S" & "</B>"
                        WSPCANTDIAS = wkproccodFC * 360
                    Case Else
                        wkproccodF = "<I>SIN DEFINIR</I>"
                        WSPCANTDIAS = 0
                End Select

                wkproccodPER = "<I>SIN EVALUACIONES</I>"
                WSPDIF = 9999
                WSPEVAL = "N"

                StrSql = ""
                StrSql = StrSql & " Select DISTINCT C.PERIODO, D.INICIO, D.PERIODO, D.ESTADO "
                StrSql = StrSql & "  From CSRELPATHDESCRI A "
                StrSql = StrSql & " INNER JOIN OBJETOS B ON (A.COBJ = B.CODOBJETO) "
                StrSql = StrSql & " INNER JOIN ORRSKCIRCUITO C ON (A.COBJ = C.CODOBJETO) "
                StrSql = StrSql & " INNER JOIN ORPERIODOS D ON (C.PERIODO = D.OBJID ) "
                StrSql = StrSql & " WHERE C.APLIC = 'OR' AND "
                StrSql = StrSql & " ( "
                StrSql = StrSql & " (A.MC1 = 0 AND A.MC2 = " & rs(0) & ") OR "
                StrSql = StrSql & " (A.MC2 = 0 AND A.MC3 = " & rs(0) & ") OR "
                StrSql = StrSql & " (A.MC3 = 0 AND A.MC4 = " & rs(0) & ") OR "
                StrSql = StrSql & " (A.MC4 = 0 AND A.MC5 = " & rs(0) & ") OR "
                StrSql = StrSql & " (A.MC5 = 0 AND A.MC6 = " & rs(0) & ") OR "
                StrSql = StrSql & " (A.MC6 = 0 AND A.MC7 = " & rs(0) & ")  "

                StrSql = StrSql & " ) "
                StrSql = StrSql & " ORDER BY D.INICIO DESC, C.PERIODO DESC "

                ''  response.write strsql & "<br>" & chr(13)
                Set RSX5 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If Not RSX5.EOF Then

                    If (RSX5(3) & "") = "OK" Then WSPEVAL = "S"


                    wkproccodPER = "<B>" & RSX5(2) & "&nbsp;-&nbsp;" & GTFECHADORSH(RSX5(1)) & "</B>"
                    WSPDIF = DHOY - RSX5(1)


                End If

                WSPDISFRES = WSPDIF - WSPCANTDIAS


                If WSPDISFRES < 20 Then

                    If WSPDISFRES < 1 Then
                        WSPBGCL = " BGCOLOR='#00DD00' "
                    Else
                        WSPBGCL = " BGCOLOR='#F7AA22' "
                    End If


                Else

                    WSPBGCL = " BGCOLOR='#DD0000' "


                End If

                WSPDISFRES = "&nbsp;"

                Celda = ""
                '''    Celda = Celda & ofv.GenCelda("","ut style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ","CENTER","2%","","","",sImageP,"si")
                Celda = Celda & ofv.GenCelda("", "UtI style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' COLSPAN=2  ", "", "50%", "", "", "", wkproccod, "si")
                Celda = Celda & ofv.GenCelda("", "Ut style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ", "", "10%", "", "", "", wkproccodF, "si")
                Celda = Celda & ofv.GenCelda("", "Ut style='color:#000000;border-style:solid;border-width:1;border-color:#EEEEEE;' " & WSPBGCL, "", "1%", "", "", "", WSPDISFRES, "si")
                Celda = Celda & ofv.GenCelda("", "UtI style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ", "", "30%", "", "", "", wkproccodPER, "si")

                Filan = Filan & ofv.GenRow("", "AAA bgcolor=" & xTrColor & " ", "", "", "25", "", "", Celda, "si")


                If TTABMAPA > 0 Then

                    StrSql = ""


                    StrSql = StrSql & " Select  A.COBJ, W.OBJID AS PERID, W.PERIODO AS PERDESCRI, W.INICIO AS PERINICIO, "
                    StrSql = StrSql & " W.PERESTADO, C.PERIODO AS RSKPERIODO,  c.ESTADO AS RSKESTADO, "
                    StrSql = StrSql & " W.codobjeto, W.descri, W.codvinobj, W.ESTADO AS SPROESTADO "
                    StrSql = StrSql & " From CSRELPATHDESCRI A "
                    StrSql = StrSql & " INNER JOIN ( "
                    StrSql = StrSql & "    select z.objid, z.periodo, z.estado AS PERESTADO, Z.INICIO, x.*  "
                    StrSql = StrSql & "    from orperiodos z, "
                    StrSql = StrSql & "    (Select    b.codobjeto, b.descri, b.codvinobj, E.ESTADO , a.tsec "
                    StrSql = StrSql & "     From CSRELPATHDESCRI A "
                    StrSql = StrSql & "     INNER JOIN OBJETOS B ON (A.COBJ = B.CODOBJETO and (CODVINOBJ IN ('VN','AI','CR'))) "
                    StrSql = StrSql & "     INNER JOIN OBJPRTY E ON (A.COBJ = E.CODOBJETO AND E.ESTADO <> 4) ) x "
                    StrSql = StrSql & "     where z.estado = 'OK' AND Z.APLIC = 'OR' "
                    StrSql = StrSql & "     ) W "
                    StrSql = StrSql & "   ON (A.COBJ = W.CODOBJETO) "
                    StrSql = StrSql & " LEFT JOIN ORRSKCIRCUITO C "
                    StrSql = StrSql & "    ON (A.COBJ = C.CODOBJETO and W.OBJID = C.PERIODO AND c.aplic = 'OR' AND c.ESTADO <> 4 ) "

                    StrSql = StrSql & " WHERE  "
                    StrSql = StrSql & " ( "
                    StrSql = StrSql & " (A.MC1 = 0 AND A.MC2 = " & rs(0) & ") OR "
                    StrSql = StrSql & " (A.MC2 = 0 AND A.MC3 = " & rs(0) & ") OR "
                    StrSql = StrSql & " (A.MC3 = 0 AND A.MC4 = " & rs(0) & ") OR "
                    StrSql = StrSql & " (A.MC4 = 0 AND A.MC5 = " & rs(0) & ") OR "
                    StrSql = StrSql & " (A.MC5 = 0 AND A.MC6 = " & rs(0) & ") OR "
                    StrSql = StrSql & " (A.MC6 = 0 AND A.MC7 = " & rs(0) & ")  "

                    StrSql = StrSql & " ) "


                    StrSql = StrSql & " ORDER BY W.INICIO DESC, W.OBJID DESC "


                    'response.write strsql & "<br>" & chr(13)
                    Set RSX5 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not RSX5.EOF Then

                        Do While Not RSX5.EOF


                            For II = 1 To TTABMAPA

                                If RSX5(1) = TABMAPA(II, 0) Then

                                    If (RSX5(5) & "") = "" Then
                                        TABMAPA(II, 3) = TABMAPA(II, 3) + 1    'PENDIENTE
                                    ElseIf (RSX5(6) & "") <> "3" Then
                                        TABMAPA(II, 4) = TABMAPA(II, 4) + 1    'EN CURSO
                                    Else
                                        TABMAPA(II, 5) = TABMAPA(II, 5) + 1    'FINALIZADO
                                    End If



                                    Exit For
                                End If

                            Next



                            RSX5.movenext
                        Loop

                        FilaX = ""

                        WSPCANTPEN = 0
                        WSPCANTCUR = 0
                        WSPCANTFIN = 0

                        Celda = ""
                        Celda = Celda & ofv.GenCelda("", "utI", "", "70%", "", "", "", "PERIODO", "si")
                        Celda = Celda & ofv.GenCelda("", "UtI", "", "10%", "", "", "", "PENDIENTES", "si")
                        Celda = Celda & ofv.GenCelda("", "UtI", "", "10%", "", "", "", "EN CURSO", "si")
                        Celda = Celda & ofv.GenCelda("", "UtI", "", "10%", "", "", "", "FINALIZADOS", "si")

                        FilaX = FilaX & ofv.GenRow("", "AAA BGCOLOR='#E0E0E0' ", "", "", "", "", "", Celda, "si")

                        If TTABMAPA > 1 Then

                            Celda = ""
                            Celda = Celda & ofv.GenCelda("", "Ut COLSPAN=4  ", "", "", "", "", "", "<U>PENDIENTES PERIODOS ANTERIORES</U>", "si")
                            FilaX = FilaX & ofv.GenRow("", "", "", "", "", "", "", Celda, "si")


                            For II = 2 To TTABMAPA


                                WSPCANTPEN = TABMAPA(II, 3) & "&nbsp;"
                                WSPCANTCUR = TABMAPA(II, 4) & "&nbsp;"
                                WSPCANTFIN = TABMAPA(II, 5) & "&nbsp;"

                                Celda = ""
                                Celda = Celda & ofv.GenCelda("", "utI", "", "70%", "", "", "", "&nbsp;-&nbsp;" & TABMAPA(II, 1), "si")
                                Celda = Celda & ofv.GenCelda("", "Ut", "RIGHT", "10%", "", "", "", WSPCANTPEN, "si")
                                Celda = Celda & ofv.GenCelda("", "Ut", "RIGHT", "10%", "", "", "", WSPCANTCUR, "si")
                                Celda = Celda & ofv.GenCelda("", "Ut", "RIGHT", "10%", "", "", "", WSPCANTFIN, "si")

                                FilaX = FilaX & ofv.GenRow("", "AAA bgcolor='#F0F0F0' ", "", "", "", "", "", Celda, "si")

                            Next



                        End If

                        II = 1

                        WSPCANTPEN = TABMAPA(II, 3) & "&nbsp;"
                        WSPCANTCUR = TABMAPA(II, 4) & "&nbsp;"
                        WSPCANTFIN = TABMAPA(II, 5) & "&nbsp;"

                        Celda = ""
                        Celda = Celda & ofv.GenCelda("", "Ut COLSPAN=4  ", "", "", "", "", "", "<U>PERIODO ACTUAL</U>", "si")
                        FilaX = FilaX & ofv.GenRow("", "", "", "", "", "", "", Celda, "si")

                        Celda = ""
                        Celda = Celda & ofv.GenCelda("", "utI", "", "70%", "", "", "", "&nbsp;-&nbsp;" & TABMAPA(II, 1), "si")
                        Celda = Celda & ofv.GenCelda("", "Ut", "RIGHT", "10%", "", "", "", WSPCANTPEN, "si")
                        Celda = Celda & ofv.GenCelda("", "Ut", "RIGHT", "10%", "", "", "", WSPCANTCUR, "si")
                        Celda = Celda & ofv.GenCelda("", "Ut", "RIGHT", "10%", "", "", "", WSPCANTFIN, "si")

                        FilaX = FilaX & ofv.GenRow("", "AAA bgcolor='#F0F0F0' ", "", "", "", "", "", Celda, "si")

                        TablaX = ofv.GenTabla("", "AAA bgcolor='#F5F5F5' style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ", "", "80%", "", "0", "", "0", "0", "", "", FilaX, "si")


                        Celda = ""
                        Celda = Celda & ofv.GenCelda("", "ut", "CENTER", "2%", "", "", "", "&nbsp;", "si")
                        Celda = Celda & ofv.GenCelda("", "Ut COLSPAN=4 ", "RIGHT", "", "", "", "", TablaX, "si")

                        Filan = Filan & ofv.GenRow("", "", "", "", "25", "", "", Celda, "si")



                    End If

                End If


                For II = 1 To TTABMAPA

                    TABMAPA(II, 3) = 0
                    TABMAPA(II, 4) = 0
                    TABMAPA(II, 5) = 0

                Next






                If xTrColor = "#FFFFFF" Then
                    xTrColor = "#F0F0F0"
                Else
                    xTrColor = "#FFFFFF"
                End If

                '''''''''''''''''''''''''''''''''

                If wkcant = 4 Then
                    wkcant = 1
                    wkleft = wkleft + 10 + 130
                    wktop = 5
                Else
                    wkcant = wkcant + 1
                    wktop = wktop + 5 + 90
                End If


                rs.movenext
            Loop
            'RESPONSE.Write(CHR(13) & CHR(13))
            'RESPONSE.Write(SMODAL)
            'RESPONSE.Write(CHR(13) & CHR(13))
            If SMODAL = "LISTA" Then

                TABLA = "<BR>" & TABLA & ofv.GenTabla("", "", "", "99%", "", "0", "", "1", "2", "", "", Filan, "si")

            End If

        End If





        Call ofv.cerrarconsulta(rs)
        Call ofv.cerrarconn(cn)

    End If
    '''''''''''''''''''''''''''

    Celda = ""






    Celda = Celda & TablaH

    Celda = Celda & TablaVH

    Celda = Celda & TABLA




    sHTML = sHTML & Celda



    sHTML = sHTML & ofv.botoneraESSX("0", "", "", "", "no", "no", "", "", "", "no", "", "no", "")


    sHTML = sHTML & "</body></html>"

End If

response.write sHTML





%>

