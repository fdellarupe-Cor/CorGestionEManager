
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = ofv.CheckUsuario()
If ok <> "Y" Then
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    sHTML = ok
Else
    Response.expires = 0

    Server.ScriptTimeout = 360
    cantregmostrarx = 6
    cantregmoverx = CDbl(request.querystring("CantRegMover"))
    vuelta = ofv.vueltaasp("../", cantregmostrarx, cantregmoverx)
    smodal = request.querystring("MODAL")
    sMODALMAIN = request.querystring("MODALMAIN")
    SCODTRACER = request.querystring("CODTRACER")
    sFECHA = request.querystring("FECHA")
    SOPCION = request.querystring("OPCION")
    ORXLS = request.Form("ORXLS")


    If smodal = "" Then smodal = "CSN"

    sHTML = ""
    If ORXLS = "SI" Then
        Response.ContentType = "application/vnd.ms-excel"
    Else

        '******************************************************************
        '* COMIENZO CONFIGURACION DEL ENCABEZADO DE LA PAGINA EL 04/05/2010
        '******************************************************************

        'sHTML = ofv.MENUHeader("","")


        Response.Write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
        Response.Write "<tr>"
        Response.Write "<td width=100% >"
        Response.Write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"), "_", " "), "600")
        Response.Write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
        Response.Write "<tr>"
        Response.Write "<td width=10% >&nbsp;</td>"
        Response.Write "<td width=90% >&nbsp;</td>"
        Response.Write "</tr></table>"

        'Encabezados

        Formulario = Chr(13) & " <TR ><TD  class=tt width=2% >Tipo</TD><TD  class=tt width=2% >Descripcion</TD></tr>" & Chr(13)



        '******************************************************************
        '* COMIENZO CONFIGURACION DEL ENCABEZADO DE LA PAGINA EL 04/05/2010
        '******************************************************************
    End If

    WKVISION = ""



    WSPSESION = "ORWHERE"




    ofv.ObtenerAtributos session("Form"), ""
    bAgregar = ofv.AAgregar
    bModificar = ofv.AModificar
    bSuprimir = ofv.AEliminar
    bConsultar = ofv.AConsultar

    Set cn = ofv.conectar(ofv.strconn2)


    Dim TABDISTRIB(100, 5)


    Dim TABRIESGOS(110, 110)

    TTABRIESGOS1 = 0
    TTABRIESGOS2 = 0

    ITABRIESGOS1 = 0
    ITABRIESGOS2 = 0

    TABRIESGOS(10, 0) = 0                  'CODIGO
    TABRIESGOS(10, 1) = "SIN ASIGNAR"       'DESCRI
    TABRIESGOS(10, 2) = 0                  'CANTIDAD
    TABRIESGOS(10, 4) = 0                  'total

    TABRIESGOS(0, 10) = 0                  'CODIGO
    TABRIESGOS(1, 10) = "SIN ASIGNAR"       'DESCRI
    TABRIESGOS(2, 10) = 0                  'CANTIDAD
    TABRIESGOS(4, 10) = 0                  'total

    StrSql = "Select * From ORTIPOEVENTO ORDER BY SECUEN, DESCRI"
    Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not rs.EOF Then
        posi = 10
        Do While Not rs.EOF
            posi = posi + 1

            TABRIESGOS(posi, 0) = rs(0)    'CODIGO
            TABRIESGOS(posi, 1) = rs(1)    'DESCRI
            TABRIESGOS(posi, 2) = 0  'CANTIDAD
            TABRIESGOS(posi, 4) = 0  'total


            TTABRIESGOS1 = posi

            rs.Movenext
        Loop



    End If



    StrSql = "Select * From ORLINNEGO1  ORDER BY DESCRIpcion"
    Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not rs.EOF Then
        posi = 10
        Do While Not rs.EOF
            posi = posi + 1

            TABRIESGOS(0, posi) = rs(0)    'CODIGO
            TABRIESGOS(1, posi) = rs(1)    'DESCRI
            TABRIESGOS(2, posi) = 0  'CANTIDAD
            TABRIESGOS(4, posi) = 0  'total


            TTABRIESGOS2 = posi


            For ii = 10 To TTABRIESGOS1
                TABRIESGOS(ii, posi) = 0  'monto
            Next

            rs.Movenext
        Loop



    End If

    TABRIESGOS(10, 10) = 0  'monto




    If SOPCION = "VS" Then


        SORCODIGO = request.Form("ORCODIGO")
        sORCARACTER = request.Form("ORCARACTER")
        SORTEXTO = request.Form("ORTEXTO")
        sORRIESGO = request.Form("ORRIESGO")
        SORCUENTA = request.Form("ORCUENTA")
        SORCLIENTE = request.Form("ORCLIENTE")
        SORMONTOD = request.Form("ORMONTOD")
        SORMONTOH = request.Form("ORMONTOH")
        SORPROCESO = request.Form("ORPROCESO")
        SORLOTE = request.Form("ORLOTE")

        SOREVALUADOS = request.Form("OREVALUADOS")
        SORRIESGOSIDEN = request.Form("ORRIESGOSIDEN")
        SORSPAIDEN = request.Form("ORSPAIDEN")

        SORFECINID = request.Form("ORFECINID")
        SORFECINIH = request.Form("ORFECINIH")

        '******************************************************************
        '* COMIENZO DE VARIABLES PEDIDAS EL 03/05/2010
        '******************************************************************

        SORFECREGCONTIN = request.Form("ORFECREGCONTIN")
        SORFECREGCONTOUT = request.Form("ORFECREGCONTOUT")
        SORFECCARGAIN = request.Form("ORFECCARGAIN")
        SORFECCARGAOUT = request.Form("ORFECCARGAOUT")
        SORFECDESCUIN = request.Form("ORFECDESCUIN")
        SORFECDESCUOUT = request.Form("ORFECDESCUOUT")
        SORFECFININ = request.Form("ORFECFININ")
        SORFECFINOUT = request.Form("ORFECFINOUT")
        SORFECESTFININ = request.Form("ORFECESTFININ")
        SORFECESTFINOUT = request.Form("ORFECESTFINOUT")
        SORESTADOS = request.Form("ORESTADOS")

        '******************************************************************
        '* FIN DE VARIABLES PEDIDAS EL 03/05/2010
        '******************************************************************


        SORCODIGOW = ""
        sORCARACTERW = ""
        SORTEXTOW = ""
        sORRIESGOW = ""
        SORCUENTAW = ""
        SORCLIENTEW = ""
        SORMONTODW = ""
        SORMONTOHW = ""
        SORPROCESOW = ""
        SORLOTEW = ""
        SOREVALUADOSW = ""
        SORRIESGOSIDENW = ""
        SORSPAIDENW = ""

        SORFECINIDW = ""
        SORFECINIHW = ""

        '******************************************************************
        '* COMIENZO DE VARIABLES PEDIDAS EL 03/05/2010
        '******************************************************************

        SORFECREGCONTINW = ""
        SORFECREGCONTOUTW = ""
        SORFECCARGAINW = ""
        SORFECCARGAOUTW = ""
        SORFECDESCUINW = ""
        SORFECDESCUOUTW = ""
        SORFECFININW = ""
        SORFECFINOUTW = ""
        SORFECESTFININW = ""
        SORFECESTFINOUTW = ""
        SORESTADOSW = ""

        '******************************************************************
        '* FIN DE VARIABLES PEDIDAS EL 03/05/2010
        '******************************************************************

        SCONECTOR = ""
        SWHERE = ""

        If smodal = "REG" Then
            SCONECTOR = " AND "
        Else
            SWHERE = " WHERE "
        End If




        session(WSPSESION) = ""


        If SORCODIGO <> "TODOS" Then
            SORCODIGOW = " (A.OBJID = '0" & SORCODIGO & "') "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORCODIGOW
            SCONECTOR = " AND "

        End If



        If SORMONTOD <> "TODOS" Then
            SORMONTODW = " (NOT A.TRCIMPBRUTO < '0" & SORMONTOD & "') "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORMONTODW
            SCONECTOR = " AND "

        End If

        If SORMONTOH <> "TODOS" Then
            SORMONTOHW = " (NOT A.TRCIMPBRUTO > '0" & SORMONTOH & "') "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORMONTOHW
            SCONECTOR = " AND "

        End If


        WWSORFECINIDDSC = ""
        WWSORFECINIHDSC = ""

        If SORFECINID = "" Then SORFECINID = "TODOS"
        If SORFECINIH = "" Then SORFECINIH = "TODOS"


        If SORFECINID <> "TODOS" Then

            If IsDate(SORFECINID) Then
                WWSORFECINID = CLng(CDate(SORFECINID) - CLng(CDate("01-01-1900")))
                WWSORFECINIDDSC = " DESDE " & CDate(SORFECINID)
            Else
                WWSORFECINID = CLng(Date) - CLng(CDate("01-01-1900"))
                WWSORFECINIDDSC = " DESDE " & Date
            End If



            SORFECINIDW = " (NOT A.FECORIG < " & WWSORFECINID & ") "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORFECINIDW
            SCONECTOR = " AND "

        End If

        If SORFECINIH <> "TODOS" Then

            If IsDate(SORFECINIH) Then
                WWSORFECINIH = CLng(CDate(SORFECINIH) - CLng(CDate("01-01-1900")))
                WWSORFECINIHDSC = " HASTA " & CDate(SORFECINIH)
            Else
                WWSORFECINIH = CLng(Date) - CLng(CDate("01-01-1900"))
                WWSORFECINIHDSC = " HASTA " & Date
            End If



            SORFECINIHW = " (NOT A.FECORIG > " & WWSORFECINIH & ") "

            session(WSPSESION) = session(WSPSESION) & SCONECTOR & SORFECINIHW
            SCONECTOR = " AND "

        End If

	'*******************************************************************************
	'* Comienzo de agregado de los nuevos parámetros de filtrado 04/05/2010
	'*******************************************************************************

		'******************************************************************************
		'* INICIO FECHA DE REGISTRACIÓN CONTABLE
		'******************************************************************************
		IF SORFECREGCONTIN = "" THEN SORFECREGCONTIN = "TODOS"
		IF SORFECREGCONTOUT = "" THEN SORFECREGCONTOUT = "TODOS"

		IF SORFECREGCONTIN <> "TODOS" THEN
		   IF ISDATE(SORFECREGCONTIN) THEN
			  WWSORFECREGCONTIN = CLNG(CDATE(SORFECREGCONTIN) - CLNG(CDATE("01-01-1900")))
			  WWSORFECREGCONTINDSC = " DESDE " & CDATE(SORFECREGCONTIN)
		   ELSE
			  WWSORFECREGCONTIN = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECREGCONTINDSC = " DESDE " & DATE()
		   END IF
		   SORFECREGCONTINW   = " (NOT A.FECCONTA < " & WWSORFECREGCONTIN & ") "
		   SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECREGCONTINW
		   SCONECTOR = " AND "
		END IF

		IF SORFECREGCONTOUT <> "TODOS" THEN
		   IF ISDATE(SORFECREGCONTOUT) THEN
			  WWSORFECREGCONTOUT = CLNG(CDATE(SORFECREGCONTOUT) - CLNG(CDATE("01-01-1900")))
			  WWSORFECREGCONTOUTDSC = " HASTA " & CDATE(SORFECREGCONTOUT)
		   ELSE
			  WWSORFECREGCONTOUT = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECREGCONTOUTDSC = " HASTA " & DATE()
		   END IF
		   SORFECREGCONTOUTW   = " (NOT A.FECCONTA > " & WWSORFECREGCONTOUT & ") "
		   SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECREGCONTOUTW
		   SCONECTOR = " AND "
		END IF

		'******************************************************************************
		'* FIN FECHA DE REGISTRACIÓN CONTABLE
		'******************************************************************************

		'******************************************************************************
		'* INICIO FECHA DE CARGA
		'******************************************************************************
		IF SORFECCARGAIN 	= "" THEN SORFECCARGAIN = "TODOS"
		IF SORFECCARGAOUT 	= "" THEN SORFECCARGAOUT = "TODOS"

		IF SORFECCARGAIN <> "TODOS" THEN
		   IF ISDATE(SORFECCARGAIN) THEN
			  WWSORFECCARGAIN 		= CLNG(CDATE(SORFECCARGAIN) - CLNG(CDATE("01-01-1900")))
			  WWSORFECCARGAINDSC 	= " DESDE " & CDATE(SORFECCARGAIN)
		   ELSE
			  WWSORFECCARGAIN 		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECCARGAINDSC 	= " DESDE " & DATE()
		   END IF
		   SORFECCARGAINW			= " (NOT A.FECREGIS < " & WWSORFECCARGAIN & ") "
		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECCARGAINW
		   SCONECTOR = " AND "
		END IF

		IF SORFECCARGAOUT <> "TODOS" THEN
		   IF ISDATE(SORFECCARGAOUT) THEN
			  WWSORFECCARGAOUT	= CLNG(CDATE(SORFECCARGAOUT) - CLNG(CDATE("01-01-1900")))
			  WWSORFECCARGAOUTDSC = " HASTA " & CDATE(SORFECCARGAOUT)
		   ELSE
			  WWSORFECCARGAOUT	= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECCARGAOUTDSC = " HASTA " & DATE()
		   END IF
		   SORFECCARGAOUTW 		= " (NOT A.FECREGIS > " & WWSORFECCARGAOUT & ") "
		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECCARGAOUTW
		   SCONECTOR 				= " AND "
		END IF

		'******************************************************************************
		'* FIN FECHA DE CARGA
		'******************************************************************************

		'******************************************************************************
		'* INICIO FECHA DE DESCUBRIMIENTO
		'******************************************************************************
		IF SORFECDESCUIN 	= "" THEN SORFECDESCUIN = "TODOS"
		IF SORFECDESCUOUT 	= "" THEN SORFECDESCUOUT = "TODOS"

		IF SORFECDESCUIN <> "TODOS" THEN
		   IF ISDATE(SORFECDESCUIN) THEN
			  WWSORFECDESCUIN		= CLNG(CDATE(SORFECDESCUIN) - CLNG(CDATE("01-01-1900")))
			  WWSORFECDESCUINDSC	= " DESDE " & CDATE(SORFECDESCUIN)
		   ELSE
			  WWSORFECDESCUIN		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECDESCUINDSC	= " DESDE " & DATE()
		   END IF
		   SORFECDESCUINW			= " (NOT A.FECIDENT < " & WWSORFECDESCUIN & ") "
		   SESSION(WSPSESION)		=  SESSION(WSPSESION) & SCONECTOR & SORFECDESCUINW
		   SCONECTOR 				= " AND "
		END IF

		IF SORFECDESCUOUT <> "TODOS" THEN
		   IF ISDATE(SORFECDESCUOUT) THEN
			  WWSORFECDESCUOUT		= CLNG(CDATE(SORFECDESCUOUT) - CLNG(CDATE("01-01-1900")))
			  WWSORFECDESCUOUTDSC 	= " HASTA " & CDATE(SORFECDESCUOUT)
		   ELSE
			  WWSORFECDESCUOUT		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECDESCUOUTDSC 	= " HASTA " & DATE()
		   END IF
		   SORFECDESCUOUTW 			= " (NOT A.FECIDENT > " & WWSORFECDESCUOUT & ") "
		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECDESCUOUTW
		   SCONECTOR 				= " AND "
		END IF

		'******************************************************************************
		'* FIN FECHA DE DESCUBRIMIENTO
		'******************************************************************************

		'******************************************************************************
		'* INICIO FECHA DE FINALIZACION
		'******************************************************************************
		IF SORFECFININ 	= "" THEN SORFECFININ = "TODOS"
		IF SORFECFINOUT = "" THEN SORFECCARGAOUT = "TODOS"

		IF SORFECFININ <> "TODOS" THEN
		   IF ISDATE(SORFECFININ) THEN
			  WWSORFECFININ 		= CLNG(CDATE(SORFECFININ) - CLNG(CDATE("01-01-1900")))
			  WWSORFECFININDSC 		= " DESDE " & CDATE(SORFECFININ)
		   ELSE
			  WWSORFECFININ 		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECFININDSC 		= " DESDE " & DATE()
		   END IF
		   SORFECFININW 			= " (NOT A.FECFINAL < " & WWSORFECFININ & ") "
		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECFININW
		   SCONECTOR 				= " AND "
		END IF

		IF SORFECFINOUT <> "TODOS" THEN
		   IF ISDATE(SORFECFINOUT) THEN
			  WWSORFECFINOUT 		= CLNG(CDATE(SORFECFINOUT) - CLNG(CDATE("01-01-1900")))
			  WWSORFECFINOUTDSC 	= " HASTA " & CDATE(SORFECFINOUT)
		   ELSE
			  WWSORFECFINOUT 		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECFINOUTDSC 	= " HASTA " & DATE()
		   END IF
		   SORFECFINOUTW 			= " (NOT A.FECFINAL > " & WWSORFECFINOUT & ") "
		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECFINOUTW
		   SCONECTOR 				= " AND "
		END IF

		'******************************************************************************
		'* FIN FECHA DE FINALIZACION
		'******************************************************************************

		'******************************************************************************
		'* INICIO FECHA ESTADO DE FINALIZACION
		'******************************************************************************
		IF SORFECESTFININ	= "" THEN SORFECESTFININ = "TODOS"
		IF SORFECESTFINOUT 	= "" THEN SORFECESTFINOUT = "TODOS"

		IF SORFECESTFININ <> "TODOS" THEN
		   IF ISDATE(SORFECESTFININ) THEN
			  WWSORFECESTFININ		= CLNG(CDATE(SORFECESTFININ) - CLNG(CDATE("01-01-1900")))
			  WWSORFECESTFININDSC 	= " DESDE " & CDATE(SORFECESTFININ)
		   ELSE
			  WWSORFECESTFININ		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECESTFININDSC 	= " DESDE " & DATE()
		   END IF
		   SORFECESTFININW 			= " (NOT A.FECFINENR < " & WWSORFECESTFININ & ") "
		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECESTFININW
		   SCONECTOR 				= " AND "
		END IF

		IF SORFECESTFINOUT <> "TODOS" THEN
		   IF ISDATE(SORFECESTFINOUT) THEN
			  WWSORFECESTFINOUT 		= CLNG(CDATE(SORFECESTFINOUT) - CLNG(CDATE("01-01-1900")))
			  WWSORFECESTFINOUTDSC 		= " HASTA " & CDATE(SORFECESTFINOUT)
		   ELSE
			  WWSORFECESTFINOUT 		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
			  WWSORFECESTFINOUTDSC 		= " HASTA " & DATE()
		   END IF
		   SORFECESTFINOUTW			= " (NOT A.FECFINENR > " & WWSORFECESTFINOUT & ") "
		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECESTFINOUTW
		   SCONECTOR 				= " AND "
		END IF

		'******************************************************************************
		'* FIN FECHA ESTADO DE FINALIZACION
		'******************************************************************************

'		'******************************************************************************
'		'* INICIO FECHA DE ESTADO FINALIZADO
'		'******************************************************************************
'		IF SORFECESTFININ 		= "" THEN SORFECESTFININ = "TODOS"
'		IF SORFECESTFINOUT 		= "" THEN SORFECCARGAOUT = "TODOS"

'		IF SORFECESTFININ <> "TODOS" THEN
'		   IF ISDATE(SORFECESTFININ) THEN
'			  WWSORFECESTFININ 		= CLNG(CDATE(SORFECESTFININ) - CLNG(CDATE("01-01-1900")))
'			  WWSORFECESTFININDSC 	= " DESDE " & CDATE(SORFECESTFININ)
'		   ELSE
'			  SORFECESTFININ 		= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
'			  WWSORFECESTFININDSC 	= " DESDE " & DATE()
'		   END IF
'		   SORFECESTFININW 			= " (NOT A.FECFINENR < " & WWSORFECESTFININ & ") "
'		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECESTFININW
'		   SCONECTOR 				= " AND "
'		END IF

'		IF SORFECESTFINOUT <> "TODOS" THEN
'		   IF ISDATE(SORFECESTFINOUT) THEN
'			  WWSORFECESTFINOUT 	= CLNG(CDATE(SORFECESTFINOUT) - CLNG(CDATE("01-01-1900")))
'			  WWSORFECESTFINOUTDSC 	= " HASTA " & CDATE(SORFECESTFINOUT)
'		   ELSE
'			  WWSORFECESTFINOUT 	= CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
'			  WWSORFECESTFINOUTDSC 	= " HASTA " & DATE()
'		   END IF
'		   SORFECESTFINOUTW 		= " (NOT A.FECFINENR > " & WWSORFECESTFINOUT & ") "
'		   SESSION(WSPSESION) 		=  SESSION(WSPSESION) & SCONECTOR & SORFECESTFINOUTW
'		   SCONECTOR 				= " AND "
'		END IF

'		'******************************************************************************
'		'* FIN FECHA DE ESTADO FINALIZADO
'		'******************************************************************************

		'******************************************************************************
		'* COMIENZO DE BUSQUEDA POR ESTADO
		'******************************************************************************

		IF SORESTADOS = "" THEN SORESTADOS = "TODOS" END IF

		IF SORESTADOS <> "TODOS" THEN
			SORESTADOSW 		= " (A.TRCESTADO = '"& SORESTADOS &"') "
			SESSION(WSPSESION)	=  SESSION(WSPSESION) & SCONECTOR & SORESTADOSW
			SCONECTOR			= " AND "
		ELSE
			SORESTADOSW 		= " (A.TRCESTADO <> 'CAN') "
			SESSION(WSPSESION)	=  SESSION(WSPSESION) & SCONECTOR & SORESTADOSW
			SCONECTOR			= " AND "
		END IF

		'******************************************************************************
		'* FIN DE BUSQUEDA POR ESTADO
		'******************************************************************************

	'*******************************************************************************
	'* Fin de agregado de los nuevos parámetros de filtrado 04/05/2010
	'*******************************************************************************


        'session(WSPSESION) = session(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '')  AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECORIG > 0 "
		session(WSPSESION) = session(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '')  AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' "




        If session(WSPSESION) <> "" Then
            session(WSPSESION) = SWHERE & session(WSPSESION)
        End If

        WSPWHERE = session(WSPSESION)


    Else

        WSPWHERE = session(WSPSESION)


    End If





    If smodal = "REG" Or smodal = "ENR" Or smodal = "ENCA" Or smodal = "CSN" Then




        StrSql = "     select A.*, B.ACTIVIDAD FROM ORTRACER A "
        StrSql = StrSql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
        StrSql = StrSql & WSPWHERE & "  ORDER BY a.objid desc"  'A.TRCESTADO desc , A.FECREGIS DESC  "
        WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

        If smodal = "CSN" Then WKACCION = " - LISTA DE EVENTOS RAIZ"



		'response.write chr(13) & chr(13)
		'response.write strsql
		'response.write chr(13) & chr(13)




        If WWSORFECINIDDSC <> "" Or WWSORFECINIHDSC <> "" Then

            celda = ""
            celda = celda & ofv.gencelda("", "ut   ", "", "100%", "", "", "", WWSORFECINIDDSC & WWSORFECINIHDSC, "si")
            Fila = ofv.genrow("", "", "", "", "", "", "", celda, "si")

            sHTML = sHTML & ofv.GenTabla("", "AAA style='border-style:SOLID;border-width:1;BORDER-COLOR:#993333;' valign='middle'  ", "", "100% ", "20", "0", "", "-1", "1", "", "", Fila, "si")
            sHTML = sHTML & "<br>"


        End If



        Fila = ""
        Fila1 = ""

        celda = ""
        celda = celda & ofv.gencelda("", "VTIg COLSPAN=10 BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ", "", "", "", "", "", session("FormN"), "si")
        Fila1 = Fila1 & ofv.genrow("", "AAA BGCOLOR='#555555' ", "", "", "25", "", "", celda, "si")

        celda = ""
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "center", "7%", "", "", "", "RAIZ", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "center", "7%", "", "", "", "CODIGO", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "center", "25%", "", "", "", "DESCRIPCION", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "CENTER", "8%", "", "", "", "MONTO BRUTO", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "CENTER", "8%", "", "", "", "OTROS MONTOS", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "CENTER", "8%", "", "", "", "PERDIDAS", "si")
		celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "CENTER", "8%", "", "", "", "NO CONTABLE", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "CENTER", "8%", "", "", "", "RECUPEROS", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "CENTER", "8%", "", "", "", "DESAFECTACIONES", "si")
        celda = celda & ofv.gencelda("", "VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ", "CENTER", "8%", "", "", "", "MONTO NETO", "si")
        Fila1 = Fila1 & ofv.genrow("", "AAA BGCOLOR='#555555' ", "", "", "25", "", "", celda, "si")


        WSPIMPOFTT = 0
        WSPCANTFTT = 0

        wkmontonetoT = 0
        wkmontoRECUT = 0
        wkmontoBRUTT = 0
        wkmontoOMONT = 0
        wkmontoPERDT = 0
        wkmontoDESAT = 0
		wkmontoNOCONT = 0


        'Encabezados

        Formulario = ""
        FILAX = ""
        celda = ""


     '    response.write strsql & chr(13)& chr(13)
        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        If rs.EOF Then


        Else
            posi = 0

            xTrColor = "#f0f0f0"



            posi = 0
            Do While (Not rs.EOF)



                wkmontonetox = rs(16) & ""

                If wkmontonetox = "" Or wkmontonetox = "" Then


                Else



                    ''''''''' BUSCA VINCULADOS '''''''''''''''''''''


                    StrSql = "     select A.*, B.ACTIVIDAD FROM ORTRACER A "
                    StrSql = StrSql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
                    StrSql = StrSql & " WHERE A.TRCCODTRACERP = '" & rs("CODTRACER") & "'  ORDER BY A.TRCESTADO desc , A.FECREGIS DESC  "

					'response.write chr(13) & chr(13)
					'response.write strsql
					'response.write chr(13) & chr(13)

					Set rsX1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If rsX1.EOF Then


                    Else
                        posi = 0

                        xTrColor = "#f0f0f0"



                        wkmontoneto = CDbl(rs(16))


                        wkmontoRECU = CDbl(rs(15))
                        wkmontoBRUT = CDbl(rs(12))
                        wkmontoOMON = CDbl(rs(14))
                        wkmontoPERD = CDbl(rs(13))
                        wkmontoDESA = CDbl(rs(36))
						if rs(42)&"" ="" then
						    wkmontoNOCON=0
						else
						   wkmontoNOCON = CDbl(  rs(42))
                        end if

						'wkmontoNOCON = CDbl( rs(42))
						'wkmontoNOCON = 0


                        wkmontonetoT = wkmontonetoT + wkmontoneto
                        wkmontoRECUT = wkmontoRECUT + wkmontoRECU
                        wkmontoBRUTT = wkmontoBRUTT + wkmontoBRUT
                        wkmontoOMONT = wkmontoOMONT + wkmontoOMON
                        wkmontoPERDT = wkmontoPERDT + wkmontoPERD
                        wkmontoDESAT = wkmontoDESAT + wkmontoDESA
						wkmontoNOCONT = wkmontoNOCONT + wkmontoNOCON



                        Celda1 = ""

                        Celda1 = Celda1 & ofv.gencelda("", "VT COLSPAN=2 BGCOLOR='#000088' style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#000000;'  ", "RIGHT", "", "", "", "", rs(0) & "&nbsp;", "si")

                        ''    Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",RS(0) & "&nbsp;","si")


                        WKMDESCRI = rs(11) & "."

                        Celda1 = Celda1 & ofv.gencelda("", "UTI style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "", "", "", "", "", WKMDESCRI, "si")

                        wspcolor = "#eeeeff"

                        WSPIMPOF = wkmontoBRUT

						IF ORXLS = "SI" THEN
							If WSPIMPOF = 0 Then
								WSPIMPOF = ""
							Else
								WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
						ELSE
							If WSPIMPOF = 0 Then
								WSPIMPOF = "&nbsp;-"
							Else
								WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
						END IF


                        WSPIMPOF = wkmontoOMON

						IF ORXLS = "SI" THEN
							If WSPIMPOF = 0 Then
								WSPIMPOF = ""
							Else

								WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF , "si")

						ELSE
							If WSPIMPOF = 0 Then
								WSPIMPOF = "&nbsp;-"
							Else

								WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
						END IF



                        WSPIMPOF = wkmontoPERD

						IF ORXLS = "SI" THEN
							If WSPIMPOF = 0 Then
								WSPIMPOF = ""
							Else

								WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")

						ELSE
							If WSPIMPOF = 0 Then
								WSPIMPOF = "&nbsp;-"
							Else

								WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
						END IF
						'***************************************************************************************
						'*
						'***************************************************************************************

							WSPIMPOF = wkmontoNOCON

						IF ORXLS = "SI" THEN
							IF WSPIMPOF = 0 THEN
							   WSPIMPOF = ""
							ELSE

							   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
							END IF

							Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF,"si")

						ELSE
							IF WSPIMPOF = 0 THEN
							   WSPIMPOF = "&nbsp;-"
							ELSE

							   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
							END IF

							Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
						END IF
						'***************************************************************************************
						'*
						'***************************************************************************************



                        WSPIMPOF = wkmontoRECU

						IF ORXLS = "SI" THEN
							If WSPIMPOF = 0 Then
								WSPIMPOF = ""
							Else
								WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF , "si")

						ELSE
							If WSPIMPOF = 0 Then
								WSPIMPOF = "&nbsp;-"
							Else
								WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
						END IF


                        WSPIMPOF = wkmontoDESA



						IF ORXLS = "SI" THEN
							If WSPIMPOF = 0 Then
								WSPIMPOF = ""
							Else

								WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")

						ELSE
							If WSPIMPOF = 0 Then
								WSPIMPOF = "&nbsp;-"
							Else

								WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
						END IF


                        WSPIMPOF = wkmontoneto

						IF ORXLS = "SI" THEN
							If WSPIMPOF = 0 Then
								WSPIMPOF = ""
							Else

								WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
						ELSE
							If WSPIMPOF = 0 Then
								WSPIMPOF = "&nbsp;-"
							Else

								WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

							End If
							Celda1 = Celda1 & ofv.gencelda("", "VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
						END IF


                        Fila1 = Fila1 & ofv.genrow("", "AAA BGCOLOR='#DDDDDD' ", "", "", "25", "", "", Celda1, "si")

                        posi = 0
                        Do While (Not rsX1.EOF)
							'response.Write(rsX1(16))

							if (rsX1(16) & "" = "") then
							wkmontoneto = 0
							else

								if(CDbl(rsX1(16)) < 0) then
								wkmontoneto = CDbl(rsX1(16))
								else
								wkmontoneto = CDbl( rsX1(16))
								end if
                            end if


                            wkmontoRECU = CDbl("0" & rsX1(15))

							if (rsX1(12) & "" = "") then
							wkmontoneto = 0
							else

								if(CDbl(rsX1(12)) < 0) then
								wkmontoneto = 0 'CDbl(rsX1(12))
								else
								wkmontoneto = 0'CDbl( rsX1(12))
								end if
                            end if


                            wkmontoOMON = CDbl("0" & rsX1(14))
                            wkmontoPERD = CDbl("0" & rsX1(13))
                            wkmontoDESA = CDbl("0" & rsX1(36))
							'wkmontoNOCON = cDBL("0" & RSx1(42))
							wkmontoNOCON = 0





                            Celda1 = ""

                            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#000088' style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#000000;'  ", "RIGHT", "", "", "", "", rs(0) & "&nbsp;", "si")


                            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#F7EEDD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "RIGHT", "", "", "", "", rsX1(0) & "&nbsp;", "si")


                            WKMDESCRI = rsX1(11) & "."

                            Celda1 = Celda1 & ofv.gencelda("", "UTI style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "", "", "", "", "", WKMDESCRI, "si")

                            wspcolor = "#eeeeff"

                            WSPIMPOF = CDbl( rsX1("trcimpbruto")) 'wkmontoBRUT



							IF ORXLS = "SI" THEN
								If WSPIMPOF = 0 Then
									WSPIMPOF = ""
								Else

									WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
							ELSE
								If WSPIMPOF = 0 Then
									WSPIMPOF = "&nbsp;-"
								Else
									WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
							END IF



                            WSPIMPOF = wkmontoOMON

							IF ORXLS = "SI" THEN
								If WSPIMPOF = 0 Then
									WSPIMPOF = ""
								Else

									WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
							ELSE
								If WSPIMPOF = 0 Then
									WSPIMPOF = "&nbsp;-"
								Else
									WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
							END IF



                            WSPIMPOF = wkmontoPERD

							IF ORXLS = "SI" THEN
								If WSPIMPOF = 0 Then
									WSPIMPOF = ""
								Else

									WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
							ELSE
								If WSPIMPOF = 0 Then
									WSPIMPOF = "&nbsp;-"
								Else
									WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
								End If
									Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
							END IF

						'***************************************************************************************
						'*
						'***************************************************************************************

							WSPIMPOF = wkmontoNOCON

							IF ORXLS = "SI" THEN
								If WSPIMPOF = 0 Then
									WSPIMPOF = ""
								Else

									WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
							ELSE
								IF WSPIMPOF = 0 THEN
								   WSPIMPOF = "&nbsp;-"
								ELSE

								   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
								END IF

								Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
							END IF
						'***************************************************************************************
						'*
						'***************************************************************************************





                            WSPIMPOF = wkmontoRECU

							IF ORXLS = "SI" THEN
								If WSPIMPOF = 0 Then
									WSPIMPOF = ""
								Else

									WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
							ELSE
								If WSPIMPOF = 0 Then
									WSPIMPOF = "&nbsp;-"
								Else
									WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)
								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
							END IF



                            WSPIMPOF = wkmontoDESA

							IF ORXLS = "SI" THEN
								If WSPIMPOF = 0 Then
									WSPIMPOF = ""
								Else

									WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")
							ELSE
								If WSPIMPOF = 0 Then
									WSPIMPOF = "&nbsp;-"
								Else

									WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
							END IF


                            WSPIMPOF = wkmontoneto

							IF ORXLS = "SI" THEN
								If WSPIMPOF = 0 Then
									WSPIMPOF = ""
								Else

									WSPIMPOF = "$ " & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF, "si")

							ELSE
								If WSPIMPOF = 0 Then
									WSPIMPOF = "&nbsp;-"
								Else

									WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

								End If
								Celda1 = Celda1 & ofv.gencelda("", "UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")
							END IF



                            Fila1 = Fila1 & ofv.genrow("", "AAA BGCOLOR='#FFFFFF' ", "", "", "25", "", "", Celda1, "si")


                            rsX1.Movenext
                        Loop

                        Celda1 = ""
                        Celda1 = Celda1 & ofv.gencelda("", "UT COLSPAN=9   ", "right", "", "", "", "", "&nbsp;", "si")

                        Fila1 = Fila1 & ofv.genrow("", "AAA  ", "", "", "", "", "", Celda1, "si")

                        '''''''''''''''''''''''''''''


                    End If

                End If



                posi = posi + 1
                rs.Movenext
            Loop




            Celda1 = ""


            WKMDESCRI = "TOTAL: "

            Celda1 = Celda1 & ofv.gencelda("", "VTI COLSPAN=3 BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "RIGHT", "", "", "", "", WKMDESCRI, "si")

            wspcolor = "#eeeeff"

            WSPIMPOF = wkmontoBRUTT


            If WSPIMPOF = 0 Then
                WSPIMPOF = "&nbsp;-"
            Else

                WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

            End If




            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")


            WSPIMPOF = wkmontoOMONT


            If WSPIMPOF = 0 Then
                WSPIMPOF = "&nbsp;-"
            Else

                WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

            End If




            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")

            WSPIMPOF = wkmontoPERDT


            If WSPIMPOF = 0 Then
                WSPIMPOF = "&nbsp;-"
            Else

                WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

            End If




            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")


			'***************************************************************************************
			'*
			'***************************************************************************************

			WSPIMPOF = wkmontoNOCONT


			IF WSPIMPOF = 0 THEN
			   WSPIMPOF = "&nbsp;-"
			ELSE
			   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
			END IF

			Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")

			'***************************************************************************************
			'*
			'***************************************************************************************

            WSPIMPOF = wkmontoRECUT


            If WSPIMPOF = 0 Then
                WSPIMPOF = "&nbsp;-"
            Else

                WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

            End If




            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")

            WSPIMPOF = wkmontoDESAT


            If WSPIMPOF = 0 Then
                WSPIMPOF = "&nbsp;-"
            Else

                WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

            End If




            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")

            WSPIMPOF = wkmontonetoT


            If WSPIMPOF = 0 Then
                WSPIMPOF = "&nbsp;-"
            Else

                WSPIMPOF = "$&nbsp;" & FormatNumber(WSPIMPOF, 2, 0, 0, -1)

            End If




            Celda1 = Celda1 & ofv.gencelda("", "VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ", "right", "", "", "", "", WSPIMPOF & "&nbsp;", "si")




            Fila1 = Fila1 & ofv.genrow("", "AAA BGCOLOR='#F0F0F0' ", "", "", "25", "", "", Celda1, "si")







            sHTML = sHTML & "<BR>"

            sHTML = sHTML & "<CENTER>" & ofv.GenTabla("", "aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ", "", "80% ", "", "0", "", "1", "1", "", "", Fila1, "si") & "</CENTER>"


        End If





    Else



    End If




    '''bgcolor='#ddddff'
    ''#F7CC88

    wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "



    'Botonera
    Agregar = ""
    Buscar = ""
    celda = ""

    sAccion = "ORMAPAEVENTOS.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER
    celda = ""    'ofv.BotonVolver(sAccion)


    sHTML = sHTML & ofv.botoneraESSX("0", "", cantregmoverx, cantregmostrarx, "no", "no", posi, "", "", "", Agregar, Buscar, celda)

    Call ofv.cerrarconsulta(rs)


    Call ofv.cerrarconn(cn)

    sHTML = sHTML & "</body></html>"
End If


Function TablaBusqueda()

    TablaBusqueda = ""

End Function
Response.Write sHTML



%>
