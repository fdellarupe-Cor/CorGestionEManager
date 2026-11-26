<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = ofv.CheckUsuario()
If ok <> "Y" Then
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    sHTML = ok
Else
    Response.expires = 0
    Set cn = ofv.conectar(ofv.strconn2)

    CantRegMostrarx = 10
    Cantregmoverx = CDbl(request.querystring("CantRegMover"))
    WSPSESION = "ORWHEREAGENDA"
	WSPSESION_PEN = "ORWHEREAGENDAPEN"
    session("filtrosql")=""
    SOPCION = request.querystring("OPCION")
    STIPO = request.form("tipo")
   ' response.write request.querystring
    IF STIPO="" THEN STIPO=request.querystring("TIPO")

    svision =  request.querystring("VISION")

    APLCHECK()


    If svision <> "" Then
        session("svision") = svision
    Else
        svision = session("svision")
    End If

    If SOPCION = "VS" Then

		'RESPONSE.Write(CHR(13) & CHR(13))
		'RESPONSE.Write("ENTRO POR VS")
		'RESPONSE.Write(CHR(13) & CHR(13))

        SORPLAN = request.Form("ORPLAN") & ""

        SORPLAN = Trim(UCase(SORPLAN))

        FINICIO = Trim(UCase(request.Form("FINICIO")))
        FFIN = Trim(UCase(request.Form("FFIN")))

        If FINICIO = "" Then FINICIO = "TODOS"
        If FFIN = "" Then FFIN = "TODOS"


        FINICIOP = Trim(UCase(request.Form("FINICIOP")))
        FFINP = Trim(UCase(request.Form("FFINP")))

        If FINICIOP = "" Then FINICIOP = "TODOS"
        If FFINP = "" Then FFINP = "TODOS"


        SORPROC = request.Form("ORPROC")
        SORPROC = Trim(UCase(SORPROC))

        'response.Write(SORPROC)
        'response.Write("??????") & chr(13)

		SESSION("PERIODOSELECCIONADO") = REQUEST.Form("ORPERIODO")

		'RESPONSE.Write(SESSION("PERIODOSELECCIONADO") & CHR(13) & REQUEST.Form("ORPERIODO"))


        If SORPROC = "NSNC" Or SORPROC = "" Then SORPROC = "TODOS"

        SORESTADO = request.Form("ORESTADO")
        SORESTADO = Trim(UCase(SORESTADO))
        If SORESTADO = "NSNC" Then SORESTADO = "TODOS"

        iF session("APLMDCLI") = "BC" then

		           SORESTADO2 = request.Form("ORESTADO02")
		           SORESTADO2 = Trim(UCase(SORESTADO2))
		           If SORESTADO2 = "NSNC" Then SORESTADO2 = "NO"

		           SORESTADO3 = request.Form("ORESTADO03")
		           SORESTADO3 = Trim(UCase(SORESTADO3))
		           If SORESTADO3 = "NSNC" Then SORESTADO3 = "NO"


		           SORESTADO4 = request.Form("ORESTADO04")
		           SORESTADO4 = Trim(UCase(SORESTADO4))
		           If SORESTADO4 = "NSNC" Then SORESTADO4 = "NO"


        END IF


        '******************************************************************
        '*
        '******************************************************************
        SORCUMPLIDO = request.Form("ORCUMPLIDO")
        SORCUMPLIDO = Trim(UCase(SORCUMPLIDO))

        'response.Write(SORCUMPLIDO)
        'response.write request.Form("ORUSUARIOS")
		IF (SORUSUARIOS = "") THEN
			SESSION("SORUSUARIOS") = request.Form("ORUSUARIOS")
		END IF


		SORUSUARIOS = SESSION("SORUSUARIOS")
        SORUSUARIOS = Trim(UCase(SORUSUARIOS))


		'response.Write(request.Form("ORUSUARIOS") & " ENTRO POR ACA")


        If (request.Form("ORSUBPROCESO") = "") Then
            SORSUBPROCESO = "TODOS"
        Else

            SORSUBPROCESO = request.Form("ORSUBPROCESO")
            SORSUBPROCESO = Trim(UCase(SORSUBPROCESO))

        End If
        'response.Write(SORSUBPROCESO) & chr(13) & chr(13)

        '******************************************************************
        '*
        '******************************************************************


        SWHERE = " WHERE "

        session(WSPSESION) = ""
		session(WSPSESION_PEN) = ""

        session("SORPROC") = SORPROC

        SORPLANW = ""
        SFINICIOW = ""
        SFFINW = ""
        SORPROCW = ""
        SORESTADOW = ""



        If SORPLAN <> "TODOS" Then
            SORPLANW = " ( F.objid = '0" & SORPLAN & "') "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORPLANW
			session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & SORPLANW
            SCONECTOR = " AND "

        End If



        If SORPROC <> "TODOS" Then
            session(WSPSESION) = ""
			session(WSPSESION_PEN) = ""
            SCONECTOR = ""
            SORPROCW = " B.APLIC = 'RG' AND  "

            SORPROCW = SORPROCW & " ("
            SORPROCW = SORPROCW & "  a.MC1 =  " & SORPROC & "   OR "
            SORPROCW = SORPROCW & "  a.MC2 =  " & SORPROC & "  OR "
            SORPROCW = SORPROCW & "  a.MC3 =  " & SORPROC & "  OR "
            SORPROCW = SORPROCW & "  a.MC4 =  " & SORPROC & "  OR "
            SORPROCW = SORPROCW & "  a.MC5 =  " & SORPROC & " OR "
            SORPROCW = SORPROCW & "  a.MC6 =  " & SORPROC & "  OR "
            SORPROCW = SORPROCW & "  a.MC7 =  " & SORPROC & "  OR "
            SORPROCW = SORPROCW & "  a.COBJ = " & SORPROC & ")"


            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORPROCW
			session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & SORPROCW
            SCONECTOR = " AND "

        End If

          'response.write SORSUBPROCESO
        If (SORSUBPROCESO <> "TODOS") Then
            SORSUBPROCESOW = (" a.cobj = " & SORSUBPROCESO)
            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORSUBPROCESOW
			session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & SORSUBPROCESOW
            SCONECTOR = " AND "
        End If


        If SORESTADO <> "TODOS" THEN
            SORESTADOW = " ( f.estado = '0" & SORESTADO & "') "




            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORESTADOW
            SCONECTOR = " AND "
       else
         IF session("APLMDCLI")="BC" THEN
	       SORESTADOW = " ( f.estado <> '5') "
               session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORESTADOW
	       SCONECTOR = " AND "
            END IF

        End If


        If FINICIO <> "TODOS" Then

            If IsDate(FINICIO) Then
                WWSORFECINIH = CLng(CDate(FINICIO)) - CLng(CDate("01-01-1900"))

            Else
                WWSORFECINIH = CLng(Date) - CLng(CDate("01-01-1900"))
            End If

            SORFECINIHW = " (Z.FTAREAS > = " & WWSORFECINIH & ") "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORFECINIHW
			session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & SORFECINIHW
            SCONECTOR = " AND "

        End If


        '******************************************************************
        '*
        '******************************************************************

        If (SORCUMPLIDO <> "TODOS" ) THEN
            If (SORCUMPLIDO = "0") Then
                WWSORCUMPLIDO = CLng(Date) - CLng(CDate("01-01-1900"))
                SSORCUMPLIDO = " (f.FECCUMPLI < " & WWSORCUMPLIDO & ") "

                session(WSPSESION) = session(WSPSESION) & SCONECTOR & SSORCUMPLIDO
                SCONECTOR = " AND "

            End If

            If (SORCUMPLIDO = "1") Then
                WWSORCUMPLIDO = CLng(Date) - CLng(CDate("01-01-1900"))
                SSORCUMPLIDO = " (f.FECCUMPLI >= " & WWSORCUMPLIDO & ") "

                session(WSPSESION) = session(WSPSESION) & SCONECTOR & SSORCUMPLIDO
                SCONECTOR = " AND "

            End If
            'RESPONSE.Write("") & CHR(13) & CHR(13)
            'RESPONSE.Write( "HOLA " & SSORCUMPLIDO & "  " & SOPCION)
            'RESPONSE.Write("") & CHR(13) & CHR(13)
        End If

        'If (SORUSUARIOS <> "TODOS") AND svision = "AGNALL" Then
        '    SSORUSUARIOS = " (f.RESPONSABLE = '" & SORUSUARIOS & "') "
        '    session(WSPSESION) = session(WSPSESION) & SCONECTOR & SSORUSUARIOS
        '    SCONECTOR = " AND "
            'RESPONSE.Write("") & CHR(13) & CHR(13)
            'RESPONSE.Write( "HOLA " & SORUSUARIOS  & "   " & WSPSESION)
            'RESPONSE.Write("") & CHR(13) & CHR(13)
        'End If

        '******************************************************************
        '*
        '******************************************************************


        If FFIN <> "TODOS" Then

            If IsDate(FFIN) Then
                WWSORFECINIH = CLng(CDate(FFIN)) - CLng(CDate("01-01-1900"))

            Else
                WWSORFECINIH = CLng(Date) - CLng(CDate("01-01-1900"))
            End If

            SORFECINIHW = " (Z.FTAREAS  < = " & WWSORFECINIH & ") "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORFECINIHW
			session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & SORFECINIHW
            SCONECTOR = " AND "

        End If


        If FINICIOP <> "TODOS" Then

            If IsDate(FINICIOP) Then
                WWSORFECINIH = CLng(CDate(FINICIOP)) - CLng(CDate("01-01-1900"))

            Else
                WWSORFECINIH = CLng(Date) - CLng(CDate("01-01-1900"))
            End If

            SORFECINIHW = " (F.FECCUMPLI > = " & WWSORFECINIH & ") "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORFECINIHW
			session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & SORFECINIHW
            SCONECTOR = " AND "

        End If


        If FFINP <> "TODOS" Then

            If IsDate(FFINP) Then
                WWSORFECINIH = CLng(CDate(FFINP)) - CLng(CDate("01-01-1900"))

            Else
                WWSORFECINIH = CLng(Date) - CLng(CDate("01-01-1900"))
            End If

            SORFECINIHW = " (F.FECCUMPLI  < = " & WWSORFECINIH & ") "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORFECINIHW
			session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & SORFECINIHW
            SCONECTOR = " AND "

        End If


         NUM_INFOTECH = request.Form("NUMINFOTECH")
		         NUM_INFOTECH = Trim(UCase(NUM_INFOTECH))

		 			'RESPONSE.WRITE NUM_INFOTECH & "      " & request.Form

		 		If NUM_INFOTECH <> "TODOS" and len(NUM_INFOTECH)>0 Then

		            NUM_INFOTECHW = " (isnull(F.NUM_INFOTECH,' ') LIKE '%" & NUM_INFOTECH & "%' ) "
		         '      NUM_INFOTECHW = " (F.NUM_INFOTECH LIKE '%" & NUM_INFOTECH & "%' ) "
		            session(WSPSESION) = session(WSPSESION) & SCONECTOR & NUM_INFOTECHW
		 			'session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & NUM_INFOTECHW
		            SCONECTOR = " AND "

        End If


        'If svision = "AGN" Then
        '    session(WSPSESION) = SWHERE & session(WSPSESION) & SCONECTOR & " (responsable='" & session("usuario") & "')"
        'Else

            If session(WSPSESION) <> "" Then
                session(WSPSESION) = SWHERE & session(WSPSESION)
            End If

			If session(WSPSESION_PEN) <> "" Then
				session(WSPSESION_PEN) = SWHERE & session(WSPSESION_PEN)
            End If

        'End If

        WSPWHERE = session(WSPSESION)
		WSPWHERE_PEN = session(WSPSESION_PEN)


    Else
        WSPWHERE = session(WSPSESION)
		WSPWHERE_PEN = session(WSPSESION_PEN)
        SORPROC = session(SORPROC)
		SORUSUARIOS = SESSION("SORUSUARIOS")

    End If



  '  response.write WSPWHERE

	    '***************************************************************************
        '* Agregado de tipos de filtros
        '***************************************************************************
        SFILTROS = session("SFILTROS")
        'Response.Write (Chr(13) & Chr(13))
        'Response.Write (session("SFILTROS"))
        'Response.Write (Chr(13) & Chr(13))
        If (SFILTROS = "") Then

            SFILTROS = "Filtros: "

            If (SORPLAN <> "TODOS") Then
                SFILTROS = SFILTROS & " / Plan " & SORPLAN
            End If

            If (FINICIO <> "TODOS") Then
                SFILTROS = SFILTROS & " / F. Inicio: " & FINICIO
            End If

            If (FFIN <> "TODOS") Then
                SFILTROS = SFILTROS & " - F. Fin: " & FFIN
            End If

            If (FINICIOP <> "TODOS") Then
                SFILTROS = SFILTROS & " / F. In. Pen: " & FINICIOP
            End If

            If (FFINP <> "TODOS") Then
                SFILTROS = SFILTROS & " - F. Fin Pen: " & FFINP
            End If

			If (SESSION("PERIODOSELECCIONADO") <> "TODOS") Then
                strsqlConsulta = "Select periodo as 'DESCRIPCION' From ORPERIODOS  where ESTADO IN ('OK','CL') and objid = " & SESSION("PERIODOSELECCIONADO")
				'RESPONSE.Write(strsqlConsulta)
                Set rsperiodo = ofv.crearconsultaEx(strsqlConsulta, cn, 1, parametros)
                VALOR = rsperiodo(0)
                Call ofv.cerrarconsulta(rsperiodo)
                SFILTROS = SFILTROS & " / Periodo: " & VALOR
            End If

            If (SORPROC <> "TODOS") Then
                strsqlConsulta = "select descri from objetos where codobjeto = " & SORPROC
                Set rsproceso = ofv.crearconsultaEx(strsqlConsulta, cn, 1, parametros)
                VALOR = rsproceso(0)
                Call ofv.cerrarconsulta(rsproceso)
                SFILTROS = SFILTROS & " / Proceso: " & VALOR
            End If

            If (SORSUBPROCESO <> "TODOS") Then
                strsqlConsulta = "select descri from objetos where codobjeto = " & SORSUBPROCESO
                Set rsproceso = ofv.crearconsultaEx(strsqlConsulta, cn, 1, parametros)
                VALOR = rsproceso(0)
                Call ofv.cerrarconsulta(rsproceso)
                SFILTROS = SFILTROS & " / SubProceso: " & VALOR
            End If

            If (SORESTADO <> "TODOS") Then
                strsqlConsulta = "SELECT descripcion FROM ORAGENDAESTADOS WHERE ESTADO = " & SORESTADO
                Set rsproceso = ofv.crearconsultaEx(strsqlConsulta, cn, 1, parametros)
                VALOR = rsproceso(0)
                Call ofv.cerrarconsulta(rsproceso)
                SFILTROS = SFILTROS & " / Estado: " & VALOR
            End If

            If (SORCUMPLIDO <> "TODOS") Then
                'RESPONSE.Write(CHR(13) & CHR(13))
                'RESPONSE.Write("HOLA JPS  " & SORCUMPLIDO)
                'RESPONSE.Write(CHR(13) & CHR(13))
                If (SORCUMPLIDO = "0") Then
                    SFILTROS = SFILTROS & " / Cumplido: VENCIDO"
                End If
                If (SORCUMPLIDO = "1") Then
                    SFILTROS = SFILTROS & " / Cumplido: SIN VENCER"
                End If
            End If

            If (SORUSUARIOS <> "TODOS" And (svision = "PENALL" Or svision = "AGNALL")) Then
                strsqlConsulta = "select distinct responsable from oragenda where responsable is not null and responsable = '" & SORUSUARIOS & "'"
                Set rsproceso = ofv.crearconsultaEx(strsqlConsulta, cn, 1, parametros)
                if not rsproceso.eof then
				  VALOR = getinfouser(rsproceso(0))
			     else
				   VALOR="&nbsp;"
                end if

                'VALOR = getinfouser(rsproceso(0))
                Call ofv.cerrarconsulta(rsproceso)
                SFILTROS = SFILTROS & " / Usuario: " & session("sgu_nombre")
            End If

            session("SFILTROS") = SFILTROS
            'RESPONSE.Write(CHR(13) & CHR(13))
            'response.Write(SESSION("SFILTROS"))
            'RESPONSE.Write(CHR(13) & CHR(13))
        Else
            SFILTROS = session("SFILTROS")
            'RESPONSE.Write(CHR(13) & CHR(13))
            'response.Write(SESSION("SFILTROS"))
            'RESPONSE.Write(CHR(13) & CHR(13))
        End If
        '***************************************************************************
        '*Fin agregado de tipos de filtro
        '***************************************************************************


    'RESPONSE.Write("") & CHR(13) & CHR(13)
    'RESPONSE.Write( "HOLA " & CHR(13) & WSPWHERE & CHR(13) & SOPCION)
    'RESPONSE.Write("") & CHR(13) & CHR(13)

    SVISIONMAIN = request.querystring("VISIONMAIN")
    sACTIVRIESGO = request.querystring("ACTIVRIESGO")
    SFECHA = request.querystring("FECHA")
    vuelta = ofv.vueltaasp("../", CantRegMostrarx, Cantregmoverx)

    If svision = ""  Then
        svision = "AGNALL"
        session("svision") = svision
    End If

    sHTML = ofv.FormHeader(session("FormN"))




        WKVISION = svision





    Select Case WKVISION


        Case "AGNALL"   '"PENALL"
'
            WKVISION = "AGNALL"
            WKVISDES = "AGENDA DE SEGUIMIENTO"
            sImage = "AGENDA"



    End Select


    wkproccodL=""
    celda = ""



    If (SFILTROS <> "Filtros: ") Then
        celda = celda & ofv.gencelda("", "", "", "90%", "", "", "", SFILTROS, "si")
        FilaX = ofv.genrow("", "", "", "", "", "", "", celda, "si")
        sHTML = sHTML & ofv.GenTabla("", "TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ", "", "100% ", "30", "0", "", "-1", "1", "", "", FilaX, "si")
        celda = ""
    End If
    wkproccodL=""
    celda = celda & ofv.gencelda("", "uti  ", "", "90%", "", "", "", "SEGUIMIENTO PLAN DE MITIGACION", "si")



    celda = celda & ofv.gencelda("", "ut  style='border-style:OUTSET;border-width:2;' ", "CENTER", "8%", "", "", "", wkproccodL, "si")

    celda = celda & ofv.gencelda("", "ut   ", "", "2%", "", "", "", "&nbsp;", "si")
    Fila = ofv.genrow("", "", "", "", "", "", "", celda, "si")
    sHTML = sHTML & ofv.GenTabla("", "TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ", "", "100% ", "30", "0", "", "-1", "1", "", "", Fila, "si")


    ' OFV.ObtenerAtributosNodo Session("Form"), "", "USUARIO", "TODOS"
    'bUSUARIO = OFV.AConsultar

    ofv.ObtenerAtributos session("Form"), ""
    bAgregar = ofv.AAgregar
    bModificar = ofv.AModificar
    bSuprimir = ofv.AEliminar
    bConsultar = ofv.AConsultar




    If session("ORPERIODO") = "" Then

        strsql = "Select * From ORPERIODOS  where ESTADO IN ('OK','CL')  "
        strsql = strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"
        Set rs = ofv.crearconsultaEx(strsql, cn, 1, parametros)
        If Not rs.EOF Then
            session("ORPERIODO") = rs(0)
            session("ORPERIODODESC") = rs(1)
        End If

    End If
    strsql=""
    'RESPONSE.WRITE(CHR(13) & CHR(13))
	'RESPONSE.WRITE svision
	'RESPONSE.WRITE(CHR(13) & CHR(13))



        WKSETTABLA = "ORAGENDA"

        xTrColor = "#FFFFFF"

    If svision = "AGNALL" Then

			'response.Write(chr(13) & chr(13))
			'response.Write("entro por agnall")
			'response.Write(chr(13) & chr(13))

	        xTrColor = "#FFFFFF"

	        wkstylo1 = " style='background-color:#F7CC88;color:#ffffff;' "

	        sImageva = "<font face='wINGdings' size=2 color='#990000'>&#0110;</font>"
	        sImagevM = "<font face='wINGdings' size=2 color='#f7aa22'>&#0110;</font>"
	        sImagevB = "<font face='wINGdings' size=2 color='#009900'>&#0110;</font>"
	        sImagevN = "<font face='wINGdings' size=2 color='#dddddd'>&#0110;</font>"

	        sImageresp = "<font face='wEBdings' size=2 color='#990000' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:hand;' >&#0128;</font>"

	        sImagePROC = "<font face='wEBdings' size=2 color='#888888' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:hand;' >2</font>"

	        sImagePROY = "<font face='wEBdings' size=2 color='#6699cc' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:hand;' >&#0204;</font>"

	        sImageFNL = "<font face='webdings' size=3 color='#009900' style='border-style:OUTSET;border-width:1;background-color:#F7eedd;' >&#097;</font>"

	        'Encabezados
	        Formulario = ""

	        celda = ""    'ofv.gencelda("","VTI style='background-color:#F7CC88;'  ","","75%","","","","PLAN","si")

	        celda = celda & ofv.gencelda("", "VTI style='background-color:#F7CC88;' COLSPAN=3 ", "", "75%", "", "", "", "PROCESO / ACTIVIDAD - RIESGO", "si")
	        celda = celda & ofv.gencelda("", "VTI style='background-color:#F7DD88;' COLSPAN=1 ", "", "10%", "", "", "", "F.CUMPLIMIENTO", "si")
	        celda = celda & ofv.gencelda("", "VTI style='background-color:#F7DD88;'  ", "", "7%", "", "", "", "T.PENDIENTES", "si")
	        celda = celda & ofv.gencelda("", "VTI style='background-color:#F7EE88;text-align: center;' colspan=2 ", "", "8%", "", "", "", "ESTADO", "si")

	        'Formulario = ofv.genrow("", "", "", "", "25", "", "", celda, "si")


	        WKSETTABLA = "ORAGENDA F "

			'response.Write(SESSION("PERIODOSELECCIONADO"))
	        Select Case SORPROC
				'NUM_INFOTECH = request.Form("NUMINFOTECH")
				'NUM_INFOTECH = Trim(UCase(NUM_INFOTECH))

				'RESPONSE.WRITE NUM_INFOTECH & "      " & request.Form

				'If NUM_INFOTECH <> "TODOS" Then

	         '   NUM_INFOTECHW = " (F.NUM_INFOTECH LIKE '%" & NUM_INFOTECH & "%' ) "

	         '   session(WSPSESION) = session(WSPSESION) & SCONECTOR & NUM_INFOTECHW
			'	session(WSPSESION_PEN) = session(WSPSESION_PEN) & SCONECTOR & NUM_INFOTECHW
	         '   SCONECTOR = " AND "

	        'End If
	            Case "TODOSxx"
				'IF (SORPROC = "TODOS") THEN
	                strsql = "Select distinct * From " & WKSETTABLA
	                strsql = strsql & " left join (Select min(fecfinal) as ftareas ,min(codagenda)as codagenda From ORAGENDADET where ESTADO='PENDIENTE' group by codagenda ) z on (F.codagenda=z.codagenda)  "


					IF (SESSION("PERIODOSELECCIONADO") = "" OR SESSION("PERIODOSELECCIONADO") = "TODOS") THEN
						if 	(WSPWHERE = "") then
							IF (SORUSUARIOS <> "TODOS" AND svision <> "AGN") THEN
								strsql = strsql & " WHERE (f.RESPONSABLE = '" & SORUSUARIOS & "')  Order By f.FECCUMPLI DESC, f.ESTADO "
							ELSE
								IF (svision = "AGN") THEN
									strsql = strsql & " WHERE (f.RESPONSABLE = '" & session("usuario") & "')  Order By f.FECCUMPLI DESC, f.ESTADO "
								ELSE
									strsql = strsql & " Order By f.FECCUMPLI DESC, f.ESTADO "
								END IF
							END IF
						ELSE
							IF (SORUSUARIOS <> "TODOS" AND svision <> "AGN") THEN
								strsql = strsql & WSPWHERE & " AND (f.RESPONSABLE = '" & SORUSUARIOS & "')  Order By f.FECCUMPLI DESC, f.ESTADO "
							ELSE
								IF (svision = "AGN") THEN
									strsql = strsql & WSPWHERE & " AND (f.RESPONSABLE = '" & session("usuario") & "')  Order By f.FECCUMPLI DESC, f.ESTADO "
								ELSE
									strsql = strsql & WSPWHERE & "  Order By f.FECCUMPLI DESC, f.ESTADO "
								END IF
							END IF
						END IF
						'RESPONSE.Write(CHR(13) & CHR(13))
	                	'RESPONSE.Write(strsql)
		                'RESPONSE.Write(CHR(13) & CHR(13))
					ELSE
						if 	(WSPWHERE = "") then
							IF (SORUSUARIOS <> "TODOS" AND svision <> "AGN") THEN
								strsql = strsql & " WHERE F.PERIODO = " & SESSION("PERIODOSELECCIONADO")  & " AND (f.RESPONSABLE = '" & SORUSUARIOS & "')  Order By f.FECCUMPLI DESC, f.ESTADO "
							ELSE
								IF (svision = "AGN") THEN
									strsql = strsql & " WHERE F.PERIODO = " & SESSION("PERIODOSELECCIONADO")  & " AND (f.RESPONSABLE = '" & session("usuario") & "')  Order By f.FECCUMPLI DESC, f.ESTADO "
								ELSE
									strsql = strsql & " WHERE F.PERIODO = " & SESSION("PERIODOSELECCIONADO")  & " Order By f.FECCUMPLI DESC, f.ESTADO "
								END IF
							END IF
						else
							IF (SORUSUARIOS <> "TODOS" AND svision <> "AGN") THEN
								strsql = strsql & WSPWHERE & " AND (f.RESPONSABLE = '" & SORUSUARIOS & "') AND F.PERIODO = " & SESSION("PERIODOSELECCIONADO")  & "  Order By f.FECCUMPLI DESC, f.ESTADO "
							ELSE
								IF (svision = "AGN") THEN
			                		strsql = strsql & WSPWHERE & " AND F.PERIODO = " & SESSION("PERIODOSELECCIONADO") & " AND (f.RESPONSABLE = '" & session("usuario") & "')  Order By f.FECCUMPLI DESC, f.ESTADO "
								ELSE
									strsql = strsql & WSPWHERE & " AND F.PERIODO = " & SESSION("PERIODOSELECCIONADO") & "  Order By f.FECCUMPLI DESC, f.ESTADO "
								END IF
							END IF
						end if
					END IF

	                'RESPONSE.Write(CHR(13) & CHR(13))
	                'RESPONSE.Write(strsql & "  - ENTRO POR EL 'TODOS' DEL CASE  " & SESSION("PERIODOSELECCIONADO"))
	                'RESPONSE.Write(CHR(13) & CHR(13))

	            Case Else
				'ELSE
                    session("filtrosql")= " select           f.*  "
	                strsqlEx = "  select           f.*,Z.FTAREAS, a.*, b.codvinobj,   "
	                strsqlEx =  strsqlEx & "        d.secuen, COALESCE (e.descri,D.DESCRIACC), case when b.codvinobj = 'AI' then ora.descri else g.descri end as descri "


	                strsql = strsql & "        from  "
	                Strsql = strsql & "          oragenda                       f  "
	                strsql = strsql & "          inner  join oractivriesgo      f2 on (F.ACTIVRIESGO = f2.CODRIESGO AND F.PERIODO = f2.PERIODO)"
	                strsql = strsql & "          inner  join objetos             b on (F2.codobjeto = b.codobjeto) "
	                strsql = strsql & "          INNER  JOIN csrelpathdescri     a on (a.cobj       = B.codobjeto)"
	                strsql = strsql & "          inner  join objlnk              c on (B.codobjeto  = c.codobjeto) "
	                strsql = strsql & "          inner  join objaccion           d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced AND d.codaccion = f.actividad) "


	                strsql = strsql & "          LEFT  join tipoaccalt           e on (d.tipoaccion = e.tipoaccionalt)"
		'estos dos left son para usar para activos informaticos o el resto, en el select se crea un case para saber si se debe de usar uno u el otro
			strsql = strsql & "          left  join orimpriesgo         g  on (f2.itemriesgo = g.objid) "
			strsql = strsql & "          left  join ORAMENAZAS         ora on (f2.itemriesgo = ora.objid) "

	                strsql = strsql & "  left join (Select min(fecfinal) as ftareas ,min(codagenda)as codagenda From ORAGENDADET where ESTADO='PENDIENTE' group by codagenda ) z on (F.codagenda=z.codagenda)"

					IF (SESSION("PERIODOSELECCIONADO") = "" OR SESSION("PERIODOSELECCIONADO") = "TODOS") THEN
						IF(WSPWHERE = "") then
							strsql = strsql & " where F.PERIODO = F.PERIODO  "
							 If svision = "AGN" Then
							     strsql = strsql & " AND F.RESPONSABLE = '" & session("usuario") & "' "
							  ELSEIF  svision = "AGNALL" AND SORUSUARIOS <> "TODOS" Then
		                          strsql = strsql & " AND F.RESPONSABLE = '" & SORUSUARIOS  & "' "
                             End If
						ELSE
							IF (SORUSUARIOS <> "TODOS" AND svision <> "AGN") THEN
								strsql = strsql & WSPWHERE & " AND (f.RESPONSABLE = '" & SORUSUARIOS & "') "
							ELSE
								strsql = strsql & WSPWHERE & " "
							END IF
						END IF
					ELSE
						IF(WSPWHERE = "") then
							strsql = strsql & " WHERE F.PERIODO = " & SESSION("PERIODOSELECCIONADO")  & " "
								 If svision = "AGN" Then
														                    strsql = strsql & " AND F.RESPONSABLE = '" & session("usuario") & "' "
														  ELSEIF  svision = "AGNALL" AND SORUSUARIOS <> "TODOS" Then
																		   strsql = strsql & " AND F.RESPONSABLE = '" & SORUSUARIOS  & "' "
                                             End If
						ELSE
							IF (SORUSUARIOS <> "TODOS" AND svision <> "AGN") THEN
								strsql = strsql & WSPWHERE & " AND (f.RESPONSABLE = '" & SORUSUARIOS & "') AND F.PERIODO = " & SESSION("PERIODOSELECCIONADO")  & " "
							ELSE
								strsql = strsql & WSPWHERE & " AND F.PERIODO = " & SESSION("PERIODOSELECCIONADO")  & " "
							END IF
						END IF
					END IF


	           '      If svision = "AGNALL" Then
	           '         strsql = strsql & " AND F.RESPONSABLE = '" & session("usuario") & "' "
	           '      End If
                     session("filtrosql")=session("filtrosql") & strsql
					 strsql = strsqlEx & strsql & " order by EXPR1,EXPR2,EXPR3,EXPR4,EXPR5,EXPR6,EXPR7, F.FECCUMPLI DESC, F.ESTADO"

	                'RESPONSE.Write(CHR(13) & CHR(13))
	              '  RESPONSE.Write(strsql & " - ENTRO POR EL ELSE DEL CASE")
	                'RESPONSE.Write(CHR(13) & CHR(13))
				'END IF
	        End Select

			'RESPONSE.WRITE chr(13) & strsql

	        'RESPONSE.WRITE chr(13) & session("filtrosql")

			Set rs = ofv.crearconsultaEx(strsql, cn, 1, parametros)
            response.write StrSql
            IF  STIPO="NO" THEN
					          SDESCRI="REASIGNAR"
					          celda = celda & ofv.gencelda("", "VTI style='background-color:#F7EE88;text-align: center;' colspan=2 ", "", "8%", "", "", "", SDESCRI, "si")
					        ELSE


					            SDESCRI="REASIGNAR&nbsp;TODOS"

					            sImagenP = "&#0232;"
							    Style_Flecha = "font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;"
							    SInput2 = "<a href='AsignarUserAGN.asp?IDUSER="&SORUSUARIOS&"&CantRegMover=" & CantRegMoverx & "&tipo=" & stipo
					   	        sInput2 = sInput2  & "' title='Asignar  Usuario ' > <font style=" & Style_Flecha & ">" & sImagenP & "</font> </a>"
					   	        SDESCRI=SDESCRI & sInput2
					   	        if not rs.eof then
					   	        celda = celda & ofv.gencelda("", "VTI style='background-color:#F7EE88;text-align: center;' colspan=2 ", "", "8%", "", "", "", SDESCRI, "si")
					   	        end if

		   END IF
		    Formulario = ofv.genrow("", "", "", "", "25", "", "", celda, "si")


	        If Not rs.EOF Then rs.Move (Cantregmoverx)
	        POSI = 0
	        Do While (Not rs.EOF) And (POSI < CantRegMostrarx)
				celda = ""

            sAccion = ""
            If bModificar Then
                sAccion = ""'"ORAGENDAEDIT.ASP?VISION=" & svision & "&ACTIVRIESGO=" & rs(0) & "&CantRegMover=" & Cantregmoverx
            End If



            'Descripcion
            If rs(0) < 10 Then
                Splan = "0" & rs(0)
            Else
                Splan = rs(0)
            End If
            SINPUT = sImageFNL & "&nbsp;" & Splan '& ofv.GenForm("FNEDIT" & sForm, "AAA STYLE='DISPLAY:NONE;' ", sAccion, "", "post", "", "", "si")
            'celda = celda & ofv.gencelda("", "UT ROWSPAN=2  title='EDITAR plan.' onclick='document.FNEDIT" & sForm & ".submit();' style='cursor:hand;' style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "CENTER", "1%", "", "", "", SINPUT, "si")
             celda = celda & ofv.gencelda("", "UT  ROWSPAN=2    ", "CENTER", "1%", "", "", "", SINPUT, "si")




                WKIPROCD = ""

                For WKI = 20 To 27

                    WKIPROC = rs(WKI) & ""
                    If WKIPROC <> "" Then
                        WKIPROCD = WKIPROCD & rs(WKI) & " / "
                    End If

                Next

                If Len(WKIPROCD) > 4 Then
                    WKIPROCDA = Left(WKIPROCD, Len(WKIPROCD) - 3)
                End If


            'Descripcion
            SINPUT = "<b><FONT COLOR='#555555' >" & WKIPROCDA & "</FONT></b><br><FONT COLOR='#883333' >" & rs(38) & "." & rs(39) & "</FONT> - " & rs(40)
            celda = celda & ofv.gencelda("", "UTI ", "", "30%", "", "", "", SINPUT, "si")



            sfunc = " onclick=""document.all.XTDX5.innerHTML = document.all.AGR" & POSI & ".innerHTML;document.all.DX5.style.display='block';"" "
            SINPUT = sImageresp
            celda = celda & ofv.gencelda("", "UT ROWSPAN=2  TITLE='Responsable del Plan.' " & sfunc & " style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "CENTER", "2%", "", "", "", SINPUT, "si")

        '    sfunc = " onclick=""document.all.XTDX5.innerHTML = document.all.AGP" & POSI & ".innerHTML;document.all.DX5.style.display='block';"" "
        '    SINPUT = sImagePROC
        '    celda = celda & ofv.gencelda("", "UT TITLE='Proceso-Actividad.' " & sfunc & " style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "CENTER", "2%", "", "", "", SINPUT, "si")

        '    SINPUT = sImagePROY
        '    sfunc = " onclick='document.AGPRY" & POSI & ".submit();' "
            '''  Celda = Celda & ofv.gencelda("","UT TITLE='Proyecto asociado.' " & sfunc & " ","CENTER","2%","","","",sInput,"si")

          '  SINPUT = "&nbsp;"
          '  celda = celda & ofv.gencelda("", "UT style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "CENTER", "2%", "", "", "", SINPUT, "si")

          '  SINPUT = "&nbsp;"
          '  celda = celda & ofv.gencelda("", "UT style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "CENTER", "2%", "", "", "", SINPUT, "si")

            'FECHA CUMPLIMIENTO



            If CLng(rs(2)) = 0 Then
                dFechan = "-------"
            Else
                dFechan = GTFechadorSh(rs(2))
            End If

            SINPUT = dFechan

            SINPUT = SINPUT

            celda = celda & ofv.gencelda("", "VTI ROWSPAN=2 style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "", "10%", "", "", "", SINPUT, "si")

            SINPUT = sImagevN

            DHOY = CLng(Date) - CLng(CDate("01-01-1900"))



            'TAREAS

            '  STRSQL="Select * From ORAGENDADET WHERE CODAGENDA = '"&RS("CODAGENDA")&"' AND ESTADO='PENDIENTE' order by fecfinal"


            '  SET RSTAREAS=ofv.crearconsultaEx(StrSql,cn,1,parametros)

            'DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))

            If (rs(5) = "0" Or rs(5) = "5") Then
                SINPUT = "-"
            Else
                If Not IsNull(rs("FTAREAS")) Then

                    DFEC = CLng(rs("FTAREAS"))

                    If DHOY > DFEC Then
                        SINPUT = sImageva
                        '     exit do
                    ElseIf DHOY = DFEC Then
                        SINPUT = sImagevM
                        '     exit do
                    Else
                        SINPUT = sImagevB
                        '     exit do
                    End If
                    '     RSTAREAS.movenext
                Else
                    SINPUT = "-"
                End If
            End If

            '   call ofv.cerrarconsulta(RSTAREAS)
            celda = celda & ofv.gencelda("", "UT ROWSPAN=2  style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;' ", "CENTER", "", "", "", "", SINPUT, "si")
            DFEC = CLng(rs(2))

            If (rs(5) = "0" Or rs(5) = "5") Then
                SINPUT = "-"
            Else
                If DHOY > DFEC Then
                    SINPUT = sImageva
                ElseIf DHOY = DFEC Then
                    SINPUT = sImagevM
                Else
                    SINPUT = sImagevB
                End If
            End If

            celda = celda & ofv.gencelda("", "UT ROWSPAN=2  style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "CENTER", "5%", "", "", "", SINPUT, "si")


            'estado

            DFEC = rs(5) & ""

            If DFEC = "0" Then
                SINPUT = "CANCELADO"
            ElseIf DFEC = "1" Then
                SINPUT = "PLANIFICADO"
            ElseIf DFEC = "2" Then
                SINPUT = "EN CURSO"
            ElseIf DFEC = "5" Then
                SINPUT = "FINALIZADO"
            Else
                SINPUT = "REPLANIFICADO"
            End If


            celda = celda & ofv.gencelda("", "UTI ROWSPAN=2 style='BORDER-BOTTOM-STYLE:solid;BORDER-BOTTOM-COLOR:#000000;BORDER-BOTTOM-WIDTH:1;'  ", "", "5%", "", "", "", SINPUT & "", "si")
             IF STIPO="NO" THEN
				sImagenP = "&#0232;"
				Style_Flecha = "font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;"
				sInput2 = "<a href='AsignarUserAGN.asp?ID=" & RS(0)& "&CantRegMover=" & CantRegMoverx & "&tipo=" & stipo
				sInput2 = sInput2  & "' title='Asignar  Usuario ' > <font style=" & Style_Flecha & ">" & sImagenP & "</font> </a>"
			   	        ELSE
			   	        sInput2="&nbsp;"
			   	        END IF
			   celda = celda & ofv.gencelda("", "UTI  ROWSPAN=2 align=center", "", "5%", "", "", "", SINPUT2, "si")

            '''''''''''''''''''''
            '' ventana de trabajo - proyectos
            '''''''''''''''''''''''''''''''''

            sAccion = ""
            If bModificar Then
                sAccion = "ORAGENDAPROY.asp?opcion=CP&AGENDA=" & rs(10) & ""
            End If

            Celdaf = Celdaf & ofv.GenForm("AGPRY" & POSI, "", sAccion, "_blank", "post", "", "", "si")


            '''''''''''''''''''''
            '' ventana de trabajo - procesos
            '''''''''''''''''''''''''''''''''



            strsql = "     select d.secuen,COALESCE (e.descri,D.DESCRIACC), "
			strsql = strsql & " 	case when b.codvinobj = 'AI' then ora.descri else g.descri end as descri,  "
			strsql = strsql & " 	a.*, b.codvinobj "
            strsql = strsql & "  from csrelpathdescri a"
            strsql = strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
            strsql = strsql & "     inner join objlnk c on (a.cobj = c.codobjeto)"
            strsql = strsql & "     inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced)"
            strsql = strsql & "     LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"
            strsql = strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad )"

			'estos dos left son para usar para activos informaticos o el resto, en el select se crea un case para saber si se debe de usar uno u el otro
			strsql = strsql & "left join orimpriesgo g on (f.itemriesgo = g.objid) "
			strsql = strsql & "left join ORAMENAZAS ora on ora.objid = f.itemriesgo "

            strsql = strsql & " where F.CODRIESGO = '" & rs(7) & "' "


            'RESPONSE.Write("") & CHR(13) & CHR(13)
            'RESPONSE.Write("3.Trabajo - Proceso: " & strsql)
            'RESPONSE.Write("") & CHR(13) & CHR(13)

            'response.write strsql & "<br>"
          ''  Set rs1 = ofv.crearconsultaEx(strsql, cn, 1, parametros)
          ''  If Not rs1.EOF Then


                WKIPROCD = ""

             ''''   For WKI = 3 To 10
           '     For WKI = 20 To 27

                 ''   WKIPROC = rs1(WKI) & ""
                   '' If WKIPROC <> "" Then
                  ''      WKIPROCD = WKIPROCD & rs1(WKI) & " / "
                  ''  End If

              '      WKIPROC = rs(WKI) & ""


             ' '      If WKIPROC <> "" Then
                  '      WKIPROCD = WKIPROCD & rs(WKI) & " / "
              '      End If

'
           ''     Next

             '   If Len(WKIPROCD) > 4 Then
             '       WKIPROCDA = Left(WKIPROCD, Len(WKIPROCD) - 3)
             '   End If

                Formulariop = ""

             '   Celdap = ""
             '   Celdap = Celdap & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "PROCESO&nbsp;", "si")
             '   SINPUT = WKIPROCDA
             '   Celdap = Celdap & ofv.gencelda("", "UtI ", "", "", "", "", "", SINPUT, "si")
             '   Formulariop = Formulariop & ofv.genrow("", "", "", "", "25", "", "", Celdap, "si")

            '    Celdap = ""
	'			IF rs("codvinobj") = "AI" THEN
	'				Celdap = Celdap & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "ACTIVO&nbsp;", "si")
	'			ELSE
	'				Celdap = Celdap & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "ACTIVIDAD&nbsp;", "si")
	'			END IF
         '       SINPUT = rs(0) & "." & rs(1)
         '       Celdap = Celdap & ofv.gencelda("", "UtI ", "", "", "", "", "", SINPUT, "si")
         '       Formulariop = Formulariop & ofv.genrow("", "", "", "", "25", "", "", Celdap, "si")

          '      Celdap = ""
	'			IF rs("codvinobj") = "AI" THEN
	'				Celdap = Celdap & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "AMENAZA&nbsp;", "si")
	'			ELSE
	' 				Celdap = Celdap & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "RIESGO&nbsp;", "si")
	'  			END IF
         '        SINPUT = rs(2)
          '      Celdap = Celdap & ofv.gencelda("", "UtI ", "", "", "", "", "", SINPUT, "si")
          '      Formulariop = Formulariop & ofv.genrow("", "", "", "", "25", "", "", Celdap, "si")


          '      Filan = ofv.GenTabla("", "aaa RULES='ROWS' STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ", "", "95%", "", "0", "", "-1", "1", "", "", Formulariop, "si")




          '      SINPUT = Filan

         '   Else

          '      SINPUT = "error al buscar la actividad"

        '    End If



         '   celda = celda & ofv.gencelda("AGP" & POSI, "aaa style='display:none;' ", "", "", "", "", "", SINPUT, "si")

            XXX = getinfouser(rs(6) & "")


            if isnumeric(session("SGU_idcostcenter")) then
             xeco=session("SGU_idcostcenter")
            else
             xeco=0
            end if

            strsql = "select  B.DESCrI FROM  ECO b where  b.ECO=0" &  xeco
          '  strsql = strsql & " WHERE A.ID = '" & rs(6) & "' "

            'RESPONSE.Write("") & CHR(13) & CHR(13)
        '    RESPONSE.Write(strsql)
            'RESPONSE.Write("") & CHR(13) & CHR(13)
            Set rs1 = ofv.crearconsultaEx(strsql, cn, 1, parametros)



            Formulariop = ""


            Celdap = ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "RESPONSABLE&nbsp;", "si")
            SINPUT = session("SGU_NOMBRE")
            Celdap = Celdap & ofv.gencelda("", "UtI ", "", "", "", "", "", SINPUT, "si")
            Formulariop = Formulariop & ofv.genrow("", "", "", "", "25", "", "", Celdap, "si")

            If Not rs1.EOF Then
                Celdap = ""
                Celdap = Celdap & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "AREA&nbsp;", "si")
                SINPUT = rs1(0)
                Celdap = Celdap & ofv.gencelda("", "UtI ", "", "", "", "", "", SINPUT, "si")
                Formulariop = Formulariop & ofv.genrow("", "", "", "", "25", "", "", Celdap, "si")

                Filan = ofv.GenTabla("", "aaa  RULES='ROWS' STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ", "", "95%", "", "0", "", "-1", "1", "", "", Formulariop, "si")
                SINPUT = Filan

            Else
                Celdap = ""
                Celdap = Celdap & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "AREA&nbsp;", "si")
                SINPUT = "SIN AREA"
                Celdap = Celdap & ofv.gencelda("", "UtI ", "", "", "", "", "", SINPUT, "si")
                Formulariop = Formulariop & ofv.genrow("", "", "", "", "25", "", "", Celdap, "si")

                Filan = ofv.GenTabla("", "aaa  RULES='ROWS' STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ", "", "95%", "", "0", "", "-1", "1", "", "", Formulariop, "si")
                SINPUT = Filan
            End If


            call ofv.cerrarconsulta(rs1)
            celda = celda & ofv.gencelda("AGR" & POSI, "aaa style='display:none;' ", "", "", "", "", "", SINPUT, "si")



            Formulario = Formulario & ofv.genrow("", "AAA bgcolor=" & xTrColor & " ", "", "", "25", "", "", celda, "si")




			Celda = ""'ofv.gencelda("", "", "", "", "", "", "", "&nbsp;", "si")

             SINPUT = "<FONT COLOR='#000080'  >" & rs(1) & "</FONT>"
             celda = celda & ofv.gencelda("", "UTI style='BORDER-STYLE:solid;BORDER-COLOR:#000088;BORDER-WIDTH:1;' ", "", "30%", "", "", "", SINPUT, "si")

             Formulario = Formulario & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")



             Celda = ofv.gencelda("","UT COLSPAN=9 style='BACKGROUND-color:#CCCCFF;' ","","","","","","&nbsp;","si")
            Formulario = Formulario & ofv.genrow("","AAA  ","","","","","",Celda,"si")

            If xTrColor = "#FFFFFF" Then
                xTrColor = "#F0F0F0"
            Else
                xTrColor = "#FFFFFF"
            End If

            POSI = POSI + 1
            rs.movenext
	        Loop
	        sHTML = sHTML & ofv.GenTabla("", "", "", "98% ", "", "0", "", "0", "0", "", "", Formulario, "si") & Celdaf



	        ''''''''''''''''''''''''''''''''''''''''''''''''
	        '' ventana de trabajo
	        '''''''''''''''''''''''''''''''''''''''''''''''

	        Filan = ""

	        sImage = "<font face='webdings' size=1 color='#FFFFFF' style='border-style:OUTSET;border-width:1;' >&#0114;</font>"

	        wkproccod2 = "<a class=nt onclick=""document.all.DX5.style.display='none';"" href='javascript:void(0);' "
	        wkproccod2 = wkproccod2 & " title='Cerrar Ventana.' >"
	        wkproccod2 = wkproccod2 & sImage
	        wkproccod2 = wkproccod2 & "</a>"

	        Formulario = ""

	        SINPUT = wkproccod2
	        celda = ""
	        celda = celda & ofv.gencelda("", "UTI" & wkstylo1, "right", "", "", "", "", SINPUT, "si")
	        Filan = ofv.genrow("", "", "", "", "15", "", "", celda, "si")
	        Formulario = Formulario & Filan

	        SINPUT = "&nbsp;"
	        celda = ""
	        celda = celda & ofv.gencelda("XTDX5", "UTI", "", "", "", "", "", SINPUT, "si")
	        Filan = ofv.genrow("", "", "", "", "", "", "", celda, "si")
	        Formulario = Formulario & Filan

	        Filan = ofv.GenTabla("", "aaa STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ", "", "95%", "", "0", "", "-1", "1", "", "", Formulario, "si")

	        Filan = "<CENTER><BR>" & Filan & "</CENTER>"

	        celda = ""
	        celda = celda & "<div id='DX5' style='z-index:1;margin-top:1;margin-left:1;position:absolute;left:400px;top:140px;display:none;"
	        celda = celda & "width:450;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
	        celda = celda & "background-color:#F9F9F9;"
	        celda = celda & "' >" & Chr(13) & Filan & "</DIV>"

        sHTML = sHTML & celda


        'Botonera
        Agregar = ""
        Buscar = ""
        celda = ""
        sAccion = "orreasignaAGENDA.asp?VISON=AGNALL"
        celda = ofv.BotonVolver(sAccion)
        WSPREEMP = "tipo=" & stipo & "&"

        sHTML = sHTML & replace(ofv.botoneraESSX("2", rs, Cantregmoverx, CantRegMostrarx, "no", "no", POSI, "", "", "no", Agregar, Buscar, celda),"cantregmostrar=",WSPREEMP & "cantregmostrar=")

       ' sHTML = sHTML & ofv.botoneraESSX("2", rs, Cantregmoverx, CantRegMostrarx, "no", "no", POSI, "", "", "no", Agregar, Buscar, celda)

        Call ofv.cerrarconsulta(rs)


        sHTML = sHTML & GTCCALENDAR("M")


    ElseIf svision = "IAG" Then




        wkstylo1 = " style='background-color:#F7CC88;color:#ffffff;' "
        Filan = ""

        strsql = "     select d.secuen,ISNULL(e.descri,D.DESCRIACC), g.descri, F.CODRIESGO, d.codaccion, f.objid,f.periodo, a.* "
        strsql = strsql & "  from csrelpathdescri a"
        strsql = strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
        strsql = strsql & "     inner join objlnk c on (a.cobj = c.codobjeto)"
        strsql = strsql & "     inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced)"
        strsql = strsql & "     LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"
        strsql = strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad )"
        strsql = strsql & "     left join orimpriesgo g on (f.itemriesgo = g.objid)"
        strsql = strsql & " where F.CODRIESGO = '" & sACTIVRIESGO & "' "

        'RESPONSE.Write("") & CHR(13) & CHR(13)
        'RESPONSE.Write(strsql)
        'RESPONSE.Write("") & CHR(13) & CHR(13)

        ' response.write strsql & "<br>"
        Set rs = ofv.crearconsultaEx(strsql, cn, 1, parametros)
        If Not rs.EOF Then

            sForm = "FORMIAG"

            strsql = "select * from oragenda "
            Set rs1 = ofv.crearconsultaEx(strsql, cn, 1, parametros)


            If SFECHA = "" Then
                SFECHA = CLng(Date) - CLng(CDate("01-01-1900"))
            End If

            sdescri = rs(0) & "." & rs(1) & " - " & rs(2)


            nCol = 70



            t3b = "<font class=at><textarea class=dt "
            t3b = t3b & "name=" & rs1(1).Name & " cols=" & nCol & " rows=2 "
            t3b = t3b & " onkeypress=" & Chr(34)
            t3b = t3b & " if (this.value.length > 195) {alert('El texto es demasiado"
            t3b = t3b & " extenso'); return false;}" & Chr(34) & " >"
            t3b = t3b & sdescri & "</textarea></font>"


            SINPUT = t3b    ' descri
            SINPUT = SINPUT & ofv.GenerarInput(rs1(2).Name, SFECHA, "HIDDEN", "", "1", "dt", "")    ' feccumpli
            SINPUT = SINPUT & ofv.GenerarInput(rs1(3).Name, "0", "HIDDEN", "", "1", "dt", "")    ' fecrevision
            SINPUT = SINPUT & ofv.GenerarInput(rs1(4).Name, "NO", "HIDDEN", "", "1", "dt", "")    ' proyrel
            SINPUT = SINPUT & ofv.GenerarInput(rs1(5).Name, "1", "HIDDEN", "", "1", "dt", "")    ' estado
            SINPUT = SINPUT & ofv.GenerarInput(rs1(6).Name, session("usuario"), "HIDDEN", "", "1", "dt", "")    ' responsable
            SINPUT = SINPUT & ofv.GenerarInput(rs1(7).Name, sACTIVRIESGO, "HIDDEN", "", "1", "dt", "")    ' activriesgo
            SINPUT = SINPUT & ofv.GenerarInput(rs1(8).Name, rs(4), "HIDDEN", "", "1", "dt", "")    ' actividad
            SINPUT = SINPUT & ofv.GenerarInput(rs1(9).Name, rs(6), "HIDDEN", "", "1", "dt", "")    ' periodo
            SINPUT = SINPUT & ofv.GenerarInput(rs1(10).Name, gtkey(rs(5)), "HIDDEN", "", "1", "dt", "")    ' codagenda
            SINPUT = SINPUT & ofv.GenerarInput(rs1(11).Name, "0", "HIDDEN", "", "1", "dt", "")    ' fecreal
            sinputd = SINPUT

            WKIPROCD = ""

            For WKI = 7 To 14

                WKIPROC = rs(WKI) & ""

                If WKIPROC <> "" Then
                    WKIPROCD = WKIPROCD & rs(WKI) & " / "
                End If


            Next

            If Len(WKIPROCD) > 4 Then
                WKIPROCD = Left(WKIPROCD, Len(WKIPROCD) - 3)
                WKIPROCDA = "<A HREF='javascript:void(0);' onclick='alert(this.innerHTML);' "
                WKIPROCDA = WKIPROCDA & " STYLE='COLOR:#888888;FONT-SIZE:7PT;FONT-FAMILY:""TIMES NEW ROMAN"";' >" & WKIPROCD & "</A>"
            End If


            celda = ""
            celda = celda & ofv.gencelda("", "VTI COLSPAN=3 " & wkstylo1, "", "", "", "", "", "PLANIFICAR MITIGACION EN AGENDA", "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")


            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "PROCESO&nbsp;", "si")
            SINPUT = WKIPROCDA
            celda = celda & ofv.gencelda("", "UtI colspan=2 ", "", "", "", "", "", SINPUT, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")


            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "ACTIVIDAD&nbsp;", "si")
            SINPUT = rs(0) & "." & rs(1)
            celda = celda & ofv.gencelda("", "UtI colspan=2 ", "", "", "", "", "", SINPUT, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")

            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "RIESGO&nbsp;", "si")
            SINPUT = rs(2)
            celda = celda & ofv.gencelda("", "vtI colspan=2 ", "", "", "", "", "", SINPUT, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")

            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "CUMPLIMIENTO&nbsp;", "si")
            SINPUT = GTFECHADOR(SFECHA)

            wkproccod2 = SINPUT

            celda = celda & ofv.gencelda("CELLFEC", "vtI colspan=2 style='cursor:hand;color:#993333;text-decoration:underline;' title='CAMBIAR FECHA DE CUMPLIMIENTO.' " & GTCSETCALR(sForm, rs1(2).Name, "N", "1", "CELLFEC", "", 0, 2, 0, 0, 0), "", "", "", "", "", wkproccod2, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")


            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "DESCRIPCION&nbsp;", "si")
            SINPUT = sinputd
            celda = celda & ofv.gencelda("", "UtI colspan=2 ", "", "", "", "", "", SINPUT, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")




            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "ESTADO&nbsp;", "si")
            SINPUT = "PLANIFICADO"
            celda = celda & ofv.gencelda("", "UtI colspan=2 ", "", "", "", "", "", SINPUT, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")

            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "RESPONSABLE&nbsp;", "si")
            SINPUT = session("SEGUSUARIO")
            celda = celda & ofv.gencelda("", "UtI colspan=2 ", "", "", "", "", "", SINPUT, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "20", "", "", celda, "si")

            SINPUT = "<INPUT TYPE='SUBMIT' CLASS=BT VALUE='AGENDAR' > "
            celda = ""
            celda = celda & ofv.gencelda("", "UTI COLSPAN=3 " & wkstylo1, "RIGHT", "", "", "", "", SINPUT, "si")
            Filan = Filan & ofv.genrow("", "", "", "", "25", "", "", celda, "si")

            '''''''''''''''''

            sAccion = "ORAGENDAEX.ASP?VISION=AAG&VISIONMAIN=" & SVISIONMAIN
            Formulario = ofv.GenForm(sForm, "", sAccion, "", "post", "", Filan, "si")

        Else

            celda = ""
            celda = celda & ofv.gencelda("", "UTI " & wkstylo1, "", "", "", "", "", "OBSERVACIONESXX&nbsp;", "si")
            SINPUT = "NO EXISTE LA ACTIVIDAD-RIESGO PARA PLANIFICAR"
            celda = celda & ofv.gencelda("", "vt colspan=2 ", "", "", "", "", "", SINPUT, "si")


            Filan = Filan & ofv.genrow("", "", "", "", "20", "", "", celda, "si")

        End If


        Filan = ofv.GenTabla("", "aaa RULES='ROWS' STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ", "", "80%", "", "0", "", "-1", "1", "", "", Formulario, "si")

        sHTML = sHTML & "<CENTER><BR><BR>" & Filan & "</CENTER>"


        'Botonera
        Agregar = ""
        Buscar = ""
        celda = ""
        sAccion = "orreasignaAGENDA.asp?VISON=AGNALL"
        celda = ofv.BotonVolver(sAccion)

        'ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "", Agregar, Buscar, celda)q

        sHTML = sHTML & GTCCALENDAR("M")

        Call ofv.cerrarconsulta(rs)

    Else



    End If


    Call ofv.cerrarconn(cn)


    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
    '''  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"

    ''    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"


    sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML



Function FechaHAB(wkfecha)


      swd = DatePart("w", CDate(wkfecha + CLng(CDate("01-01-1900"))))

      If swd > 6 Then
         wkfecha = wkfecha + 2
      End If

      If swd = 1 Then
         wkfecha = wkfecha + 1
      End If


      FechaHAB = wkfecha

End Function




%>
