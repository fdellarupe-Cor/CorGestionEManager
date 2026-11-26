<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Server.ScriptTimeout=360
  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sMODAL       = request.querystring("MODAL")
  sMODALMAIN   = request.querystring("MODALMAIN")
  sCODTRACER  = request.querystring("CODTRACER")
  sFECHA        = request.querystring("FECHA")
  sOPCION        = request.querystring("OPCION")
  ORXLS=REQUEST.FORM("ORXLS")
  if sMODAL = "" then sMODAL = "CSN"

  sHTML=""
  IF ORXLS="SI" THEN
     Response.ContentType = "application/vnd.ms-excel"
  ELSE
	'******************************************************************
	'* COMIENZO CONFIGURACION DEL ENCABEZADO DE LA PAGINA EL 04/05/2010
	'******************************************************************

	'sHTML = ofv.MENUHeader("","")
  
	response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
	response.write "<tr>"
	response.write "<td width=100% >"
	response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
	response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
	response.write "<tr>"
	response.write "<td width=10% >&nbsp;</td>"
	response.write "<td width=90% >&nbsp;</td>"
	response.write "</tr></table>"

	'Encabezados

	Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Tipo</TD><TD  class=tt width=2% >Descripcion</TD></tr>" & chr(13)
	
	'******************************************************************
	'* COMIENZO CONFIGURACION DEL ENCABEZADO DE LA PAGINA EL 04/05/2010
	'******************************************************************
  END IF

  WKVISION = ""



  WSPSESION = "ORWHERE"




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  DIM TABDISTRIB(100,5)


  DIM TABRIESGOS(110,110)

  TTABRIESGOS1 = 0
  TTABRIESGOS2 = 0

  ITABRIESGOS1 = 0
  ITABRIESGOS2 = 0

          TABRIESGOS(10,0) = 0             'CODIGO
          TABRIESGOS(10,1) = "SIN ASIGNAR" 'DESCRI
          TABRIESGOS(10,2) = 0             'CANTIDAD
          TABRIESGOS(10,4) = 0             'total

          TABRIESGOS(0,10) = 0             'CODIGO
          TABRIESGOS(1,10) = "SIN ASIGNAR" 'DESCRI
          TABRIESGOS(2,10) = 0             'CANTIDAD
          TABRIESGOS(4,10) = 0             'total

  Strsql = "Select * From ORTIPOEVENTO ORDER BY SECUEN, DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 10
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRIESGOS(POSI,1) = RS(1) 'DESCRI
          TABRIESGOS(POSI,2) = 0     'CANTIDAD
          TABRIESGOS(POSI,4) = 0     'total


          TTABRIESGOS1 = POSI

        RS.MOVENEXT
     LOOP



  END IF



  Strsql = "Select * From ORLINNEGO1  ORDER BY DESCRIpcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 10
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOS(0,POSI) = RS(0) 'CODIGO
          TABRIESGOS(1,POSI) = RS(1) 'DESCRI
          TABRIESGOS(2,POSI) = 0     'CANTIDAD
          TABRIESGOS(4,POSI) = 0     'total


          TTABRIESGOS2 = POSI


          for ii=10 to  TTABRIESGOS1
              TABRIESGOS(ii,POSI) = 0     'monto
          next

        RS.MOVENEXT
     LOOP



  END IF

  TABRIESGOS(10,10) = 0     'monto




  IF SOPCION = "VS" THEN


    SORCODIGO   = request.FORM("ORCODIGO")
    sORCARACTER = request.FORM("ORCARACTER")
    SORTEXTO    = request.FORM("ORTEXTO")
    sORRIESGO   = request.FORM("ORRIESGO")
    SORCUENTA   = request.FORM("ORCUENTA")
    SORCLIENTE  = request.FORM("ORCLIENTE")
    SORMONTOD   = request.FORM("ORMONTOD")
    SORMONTOH   = request.FORM("ORMONTOH")
    SORPROCESO  = request.FORM("ORPROCESO")
    SORLOTE     = request.FORM("ORLOTE")

    SOREVALUADOS       = request.FORM("OREVALUADOS")
    SORRIESGOSIDEN     = request.FORM("ORRIESGOSIDEN")
    SORSPAIDEN         = request.FORM("ORSPAIDEN")

    SORFECINID   = request.FORM("ORFECINID")
    SORFECINIH   = request.FORM("ORFECINIH")

	'******************************************************************
	'* COMIENZO DE VARIABLES PEDIDAS EL 03/05/2010
	'******************************************************************
	
	SORFECREGCONTIN	= request.FORM("ORFECREGCONTIN")
	SORFECREGCONTOUT= request.FORM("ORFECREGCONTOUT")
	SORFECCARGAIN	= request.FORM("ORFECCARGAIN")
	SORFECCARGAOUT	= request.FORM("ORFECCARGAOUT")
	SORFECDESCUIN	= request.FORM("ORFECDESCUIN")
	SORFECDESCUOUT	= request.FORM("ORFECDESCUOUT")
	SORFECFININ		= request.FORM("ORFECFININ")
	SORFECFINOUT 	= request.FORM("ORFECFINOUT")
	SORFECESTFININ	= request.FORM("ORFECESTFININ")
	SORFECESTFINOUT	= request.FORM("ORFECESTFINOUT")
	SORESTADOS		= request.FORM("ORESTADOS")
	SFECHASELECCION	= request.FORM("FECHASELECCION")

	'******************************************************************
	'* FIN DE VARIABLES PEDIDAS EL 03/05/2010
	'******************************************************************


    SORCODIGOW   = ""
    sORCARACTERW = ""
    SORTEXTOW    = ""
    sORRIESGOW   = ""
    SORCUENTAW   = ""
    SORCLIENTEW  = ""
    SORMONTODW   = ""
    SORMONTOHW   = ""
    SORPROCESOW  = ""
    SORLOTEW     = ""
    SOREVALUADOSW       = ""
    SORRIESGOSIDENW     = ""
    SORSPAIDENW         = ""

    SORFECINIDW   = ""
    SORFECINIHW   = ""

	'******************************************************************
	'* COMIENZO DE VARIABLES PEDIDAS EL 03/05/2010
	'******************************************************************
	
	SORFECREGCONTINW	= ""
	SORFECREGCONTOUTW	= ""
	SORFECCARGAINW		= ""
	SORFECCARGAOUTW		= ""
	SORFECDESCUINW		= ""
	SORFECDESCUOUTW		= ""
	SORFECFININW		= ""
	SORFECFINOUTW 		= ""
	SORFECESTFININW		= ""
	SORFECESTFINOUTW	= ""
	SORESTADOSW			= ""

	
	'******************************************************************
	'* FIN DE VARIABLES PEDIDAS EL 03/05/2010
	'******************************************************************

    SCONECTOR = ""
    SWHERE    = ""

    if sMODAL = "REG"  then
       SCONECTOR = " AND "
    ELSE
       SWHERE    = " WHERE "
    END IF




       SESSION(WSPSESION) =  ""


    IF SORMONTOD <> "TODOS" THEN
       SORMONTODW   = " (NOT A.TRCIMPBRUTO < '0" & SORMONTOD & "') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORMONTODW
       SCONECTOR = " AND "

    END IF

    IF SORMONTOH <> "TODOS" THEN
       SORMONTOHW   = " (NOT A.TRCIMPBRUTO > '0" & SORMONTOH & "') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORMONTOHW
       SCONECTOR = " AND "

    END IF


       WWSORFECINIDDSC = ""
       WWSORFECINIHDSC = ""

    IF SORFECINID = "" THEN SORFECINID = "TODOS"
    IF SORFECINIH = "" THEN SORFECINIH = "TODOS"


    IF SORFECINID <> "TODOS" THEN

       if IsDate(SORFECINID) then
          WWSORFECINID = CLNG(CDATE(SORFECINID) - CLNG(CDATE("01-01-1900")))
          WWSORFECINIDDSC = " DESDE " & CDATE(SORFECINID)
       else
          WWSORFECINID = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
          WWSORFECINIDDSC = " DESDE " & DATE()
       end if



       SORFECINIDW   = " (NOT A.FECORIG < " & WWSORFECINID & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIDW
       SCONECTOR = " AND "

    END IF

    IF SORFECINIH <> "TODOS" THEN

       if IsDate(SORFECINIH) then
          WWSORFECINIH = CLNG(CDATE(SORFECINIH) - CLNG(CDATE("01-01-1900")))
          WWSORFECINIHDSC = " HASTA " & CDATE(SORFECINIH)
       else
          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
          WWSORFECINIHDSC = " HASTA " & DATE()
       end if



       SORFECINIHW   = " (NOT A.FECORIG > " & WWSORFECINIH & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
       SCONECTOR = " AND "

    END IF

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

	IF (SFECHASELECCION = 1) THEN
		SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECORIG > 0 "
	ELSEIF (SFECHASELECCION = 2) THEN
		SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECCONTA > 0 "
	ELSEIF (SFECHASELECCION = 3) THEN
		SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECREGIS > 0 "
	ELSEIF (SFECHASELECCION = 4) THEN
		SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECIDENT > 0 "
	ELSEIF (SFECHASELECCION = 5) THEN
		SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECFINAL > 0 "
	ELSEIF (SFECHASELECCION = 6) THEN
		SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECFINENR > 0 "
	END IF
    
	'SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' "




    IF SESSION(WSPSESION) <> "" THEN
       SESSION(WSPSESION) =  SWHERE & SESSION(WSPSESION)
    END IF

    WSPWHERE = SESSION(WSPSESION)


  ELSE

    WSPWHERE = SESSION(WSPSESION)


  END IF





  if sMODAL = "REG" or sMODAL = "ENR"  or sMODAL = "ENCA"  or sMODAL = "CSN"  then


        WKFFECMAX = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
        WKFFECMIN = WKFFECMAX - 360

	IF (SFECHASELECCION = 1) THEN
    	Strsql = "     select MIN(A.FECORIG) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 2) THEN
		Strsql = "     select MIN(A.FECCONTA) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 3) THEN
		Strsql = "     select MIN(A.FECREGIS) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 4) THEN
		Strsql = "     select MIN(A.FECIDENT) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 5) THEN
		Strsql = "     select MIN(A.FECFINAL) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 6) THEN
		Strsql = "     select MIN(A.FECFINENR) FROM ORTRACER A "
	END IF
     Strsql = Strsql & WSPWHERE
	 
	 'RESPONSE.Write(CHR(13) & CHR(13))
	' 	 RESPONSE.Write(STRSQL)
	' RESPONSE.Write(CHR(13) & CHR(13))
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 
     if NOT rs.eof then
        if (RS(0) & "") <> "" then WKFFECMIN = RS(0)
     END IF

 '' response.write strsql & "<br>"

	IF (SFECHASELECCION = 1) THEN
    	Strsql = "     select MAX(A.FECORIG) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 2) THEN
		Strsql = "     select MAX(A.FECCONTA) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 3) THEN
		Strsql = "     select MAX(A.FECREGIS) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 4) THEN
		Strsql = "     select MAX(A.FECIDENT) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 5) THEN
		Strsql = "     select MAX(A.FECFINAL) FROM ORTRACER A "
	ELSEIF (SFECHASELECCION = 6) THEN
		Strsql = "     select MAX(A.FECFINENR) FROM ORTRACER A "
	END IF
	Strsql = Strsql & WSPWHERE
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rs.eof then
        if (RS(0) & "") <> "" then WKFFECMAX = RS(0)
     END IF
 '' response.write strsql & "<br>"


     WKFDATEMIN = CDATE(WKFFECMIN + CLNG(CDATE("01-01-1900")) )

     gtcsanio   = datepart("yyyy",WKFDATEMIN)
     gtcsmes    = datepart("m",WKFDATEMIN)

     WKFDESDE =  (gtcsanio * 100) + gtcsmes
     WKFDESDEY =  gtcsanio
     WKFDESDEM =  gtcsmes

     WKFDATEMAX = CDATE(WKFFECMAX + CLNG(CDATE("01-01-1900")) )

     gtcsanio   = datepart("yyyy",WKFDATEMAX)
     gtcsmes    = datepart("m",WKFDATEMAX)

     WKFHASTA =  (gtcsanio * 100) + gtcsmes
     WKFHASTAY =  gtcsanio
     WKFHASTAM =  gtcsmes

     IF WKFDESDEY = WKFHASTAY THEN
        WKFMESES = WKFHASTA - WKFDESDE + 1

     ELSE
        WKFMESES = WKFHASTAY - WKFDESDEY
        IF  WKFMESES > 1 THEN
            WKFMESES = (WKFMESES - 1) * 12
        ELSE
            WKFMESES = 0
        END IF

        WKFMESES = WKFMESES + ( (WKFHASTAM + 12 - WKFDESDEM) + 1)

     END IF


     DIM TABACCMES(200,420,10)

     WINITY = WKFDESDEY
     WINITM = WKFDESDEM

     FOR II=1 TO (WKFMESES + 2)

         TABACCMES(0,II,1) = WINITY
         TABACCMES(0,II,2) = WINITM

         TABACCMES(0,II,3) = 0
         TABACCMES(0,II,4) = 0

         WINITM = WINITM + 1

         IF WINITM > 12 THEN
            WINITM = 1
            WINITY = WINITY + 1
         END IF


     NEXT


     TOPEMESES = WKFMESES

     CTAINDEX = 1
     TABACCMES(CTAINDEX,0,0) = "0"
     TABACCMES(CTAINDEX,0,1) = "SIN ASIGNAR"
     FOR II=0 TO TOPEMESES
         TABACCMES(CTAINDEX,II,3) = 0
         TABACCMES(CTAINDEX,II,4) = 0
     NEXT



     Strsql = "     select A.*, B.ACTIVIDAD FROM ORTRACER A "
     Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
    
	IF (SFECHASELECCION = 1) THEN
    	Strsql = Strsql & WSPWHERE & "  ORDER BY A.trcctacble, A.FECORIG  "
	ELSEIF (SFECHASELECCION = 2) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.trcctacble, A.FECCONTA  "
	ELSEIF (SFECHASELECCION = 3) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.trcctacble, A.FECREGIS  "
	ELSEIF (SFECHASELECCION = 4) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.trcctacble, A.FECIDENT  "
	ELSEIF (SFECHASELECCION = 5) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.trcctacble, A.FECFINAL  "
	ELSEIF (SFECHASELECCION = 6) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.trcctacble, A.FECFINENR  "
	END IF
'''CTACBLE RS(29)



     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

     if sMODAL = "CSN"  then      WKACCION = " - PERDIDAS POR RUBRO CONTABLE"



 '' response.write strsql & "<br>"





   IF WWSORFECINIDDSC <> "" OR WWSORFECINIHDSC <> "" THEN

      Celda = ""
      Celda = Celda & ofv.GenCelda("","ut   ","","100%","","","",WWSORFECINIDDSC & WWSORFECINIHDSC,"si")
      Fila  = ofv.GenRow("","","","","","","",Celda,"si")

      sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;BORDER-COLOR:#993333;' valign='middle'  ","","100% ","20","0","","-1","1","","",Fila,"si")
      sHTML = sHTML & "<br>"
   END IF



  'Encabezados

   Formulario = ""
   FILAX = ""
   Celda = ""


''' encabezados

    Fila = ""
    Fila1 = ""




    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=" & ((WKFMESES + 1) * 2) + 1 & "  BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","","PERDIDAS POR RUBRO CONTABLE","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")



    Celda2 = ""
    Celda2 = Celda2 & ofv.gencelda("","VT style='COLOR:#FFFFFF;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='#990000' ","CENTER","60%","","","","MESES","si")

    Celda3 = ""
    Celda3 = Celda3 & ofv.gencelda("","VT style='COLOR:#FFFFFF;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='#888888' ","CENTER","","","","","RUBROS CONTABLES","si")


    FOR II=1 TO WKFMESES

        WSPIMPOF = TABACCMES(0,II,1) & "<br>" & RIGHT(CSTR(TABACCMES(0,II,2) + 100),2)
        Celda2 = Celda2 & ofv.gencelda("","uT COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","CENTER","","","","",WSPIMPOF,"si")

        Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#990000;'  ","CENTER","2%","","","","MONTO","si")
        Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#990000;'  ","CENTER","2%","","","","EVENTOS","si")

    NEXT

        WSPIMPOF = "TOTAL RUBRO"
        Celda2 = Celda2 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","CENTER","","","","",WSPIMPOF,"si")

        Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#990000;'  ","CENTER","2%","","","","MONTO","si")
        Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#990000;'  ","CENTER","2%","","","","EVENTOS","si")


        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda2,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#DDDDDD' ","","","25","","",Celda3,"si")


'RESPONSE.Write(CHR(13) & CHR(13))
'RESPONSE.Write(STRSQL)
'RESPONSE.Write(CHR(13) & CHR(13))

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


  ELSE
  posI = 0

  xTrColor = "#f0f0f0"

  CTAANT = "NO"

  CTATOPE  = 1
  CTAINDEX = 1

  posi = 0
  Do While (Not rs.EOF)
     wkmontonetox =  rs(16) & ""
     if wkmontonetox =  "" OR wkmontonetox =  "0" then
     else
        CTAEVENTO = RS(29) & ""
        IF CTAEVENTO <> CTAANT THEN
           IF CTAANT <> "NO" AND CTAINDEX > 1 THEN

              WSPCUENTA = TABACCMES(CTAINDEX,0,0) & "&nbsp;-&nbsp;" & TABACCMES(CTAINDEX,0,1)
              Celda3 = ""
              Celda3 = Celda3 & ofv.gencelda("","UTI BGCOLOR='#DDDDFF' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",WSPCUENTA,"si")
              WSPIMPOTT = 0
              WSPCANTTT = 0

              FOR WKFindex=1 TO WKFMESES

                  WSPIMPOC  = TABACCMES(CTAINDEX,WKFindex,3)
                  WSPCANTC  = TABACCMES(CTAINDEX,WKFindex,4)
                  WSPIMPOTT = WSPIMPOTT + WSPIMPOC
                  WSPCANTTT = WSPCANTTT + WSPCANTC

                  IF WSPIMPOC = 0 THEN
                     WSPIMPOC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOC = "$&nbsp;" & Formatnumber(WSPIMPOC ,2,0,0,-1) & "&nbsp;"
                  END IF

                  IF WSPCANTC = 0 THEN
                     WSPCANTC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTC = WSPCANTC & "&nbsp;"
                  END IF

                  Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","RIGHT","","","","",WSPIMPOC,"si")
                  Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","RIGHT","","","","",WSPCANTC,"si")

              NEXT

                  IF WSPIMPOTT = 0 THEN
                     WSPIMPOTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOTT = "$&nbsp;" & Formatnumber(WSPIMPOTT ,2,0,0,-1) & "&nbsp;"
                  END IF



                  IF WSPCANTTT = 0 THEN
                     WSPCANTTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTTT = WSPCANTTT & "&nbsp;"
                  END IF


              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPIMPOTT,"si")
              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPCANTTT,"si")

              Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#FFFFFF' ","","","25","","",Celda3,"si")




           END IF


           strsql = "Select descri From OBJCUENTA where NUMERO = '" & CTAEVENTO & "' "
		   'RESPONSE.Write(CHR(13) & CHR(13))
		   'RESPONSE.Write(STRSQL)
		   'RESPONSE.Write(CHR(13) & CHR(13))
           set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsniv.eof then
              sInputx = rsniv(0)
           ELSE
              sInputx = "NO"
           END IF
           call ofv.cerrarconsulta(rsniv)

           CTAANT = CTAEVENTO
           IF sInputx = "NO" THEN
              CTAINDEX = 1

           ELSE
              CTATOPE  = CTATOPE + 1
              CTAINDEX = CTATOPE
              TABACCMES(CTAINDEX,0,0) = CTAEVENTO
              TABACCMES(CTAINDEX,0,1) = sInputx
              FOR II=0 TO TOPEMESES
                  TABACCMES(CTAINDEX,II,3) = 0
                  TABACCMES(CTAINDEX,II,4) = 0
              NEXT

           END IF


        END IF



        wkmontoneto =  cdbl(rs(16))


'        WKFDATEMAX = CDATE(rs(3) + CLNG(CDATE("01-01-1900")) )
	IF (SFECHASELECCION = 1) THEN
		WKFDATEMAX = CDATE(rs(3) + CLNG(CDATE("01-01-1900")) )
	ELSEIF (SFECHASELECCION = 2) THEN
		WKFDATEMAX = CDATE(rs(5) + CLNG(CDATE("01-01-1900")) )
	ELSEIF (SFECHASELECCION = 3) THEN
		WKFDATEMAX = CDATE(rs(2) + CLNG(CDATE("01-01-1900")) )
	ELSEIF (SFECHASELECCION = 4) THEN
		WKFDATEMAX = CDATE(rs(4) + CLNG(CDATE("01-01-1900")) )
	ELSEIF (SFECHASELECCION = 5) THEN
		WKFDATEMAX = CDATE(rs(23) + CLNG(CDATE("01-01-1900")) )
	ELSEIF (SFECHASELECCION = 6) THEN
		WKFDATEMAX = CDATE(rs(35) + CLNG(CDATE("01-01-1900")) )
	END IF


        gtcsanio   = datepart("yyyy",WKFDATEMAX)
        gtcsmes    = datepart("m",WKFDATEMAX)


        IF WKFDESDEY = gtcsanio THEN
           WKFindex = gtcsmes - WKFDESDEM + 1

        ELSE
           WKFindex = gtcsanio - WKFDESDEY
           IF  WKFindex > 1 THEN
               WKFindex = (WKFindex - 1) * 12
           ELSE
               WKFindex = 0
           END IF

           WKFindex = WKFindex + ( (gtcsmes + 12 - WKFDESDEM) + 1)


         END IF
       '  response.write " wi:" &   WKFindex & " wm:" & wkfmeses & " DS:" & WKFDESDEY & " HS:" & gtcsanio & gtcsmes & "<br>"
         TABACCMES(CTAINDEX,WKFindex,3) = TABACCMES(CTAINDEX,WKFindex,3) + wkmontoneto
         TABACCMES(CTAINDEX,WKFindex,4) = TABACCMES(CTAINDEX,WKFindex,4) + 1

         TABACCMES(CTAINDEX,0,3) = TABACCMES(CTAINDEX,0,3) + wkmontoneto
         TABACCMES(CTAINDEX,0,4) = TABACCMES(CTAINDEX,0,4) + 1

         TABACCMES(0,WKFindex,3) = TABACCMES(0,WKFindex,3) + wkmontoneto
         TABACCMES(0,WKFindex,4) = TABACCMES(0,WKFindex,4) + 1



        wktipoevento = rs(30)
        wkdistrib    = rs(37)
        wklineaN2    = rs(8)


     end if





    posi = posi + 1
    rs.Movenext
  Loop

           IF CTAANT <> "NO" AND CTAINDEX > 1 THEN


              WSPCUENTA = TABACCMES(CTAINDEX,0,0) & "&nbsp;-&nbsp;" & TABACCMES(CTAINDEX,0,1)

              Celda3 = ""
              Celda3 = Celda3 & ofv.gencelda("","UTI BGCOLOR='#DDDDFF' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#888888;'  ","","20%","","","",WSPCUENTA,"si")

                  WSPIMPOTT = 0
                  WSPCANTTT = 0

              FOR WKFindex=1 TO WKFMESES

                  WSPIMPOC  = TABACCMES(CTAINDEX,WKFindex,3)
                  WSPCANTC  = TABACCMES(CTAINDEX,WKFindex,4)

                  WSPIMPOTT = WSPIMPOTT + WSPIMPOC
                  WSPCANTTT = WSPCANTTT + WSPCANTC


                  IF WSPIMPOC = 0 THEN
                     WSPIMPOC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOC = "$&nbsp;" & Formatnumber(WSPIMPOC ,2,0,0,-1) & "&nbsp;"
                  END IF



                  IF WSPCANTC = 0 THEN
                     WSPCANTC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTC = WSPCANTC & "&nbsp;"
                  END IF

                  Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","RIGHT","","","","",WSPIMPOC,"si")
                  Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","RIGHT","","","","",WSPCANTC,"si")

              NEXT

                  IF WSPIMPOTT = 0 THEN
                     WSPIMPOTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOTT = "$&nbsp;" & Formatnumber(WSPIMPOTT ,2,0,0,-1) & "&nbsp;"
                  END IF



                  IF WSPCANTTT = 0 THEN
                     WSPCANTTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTTT = WSPCANTTT & "&nbsp;"
                  END IF


              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPIMPOTT,"si")
              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPCANTTT,"si")

              Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#FFFFFF' ","","","25","","",Celda3,"si")




           END IF


  CTAINDEX = 1

  IF (TABACCMES(CTAINDEX,0,3) > 0) THEN


              WSPCUENTA = TABACCMES(CTAINDEX,0,0) & "&nbsp;" & TABACCMES(CTAINDEX,0,1)

              Celda3 = ""
              Celda3 = Celda3 & ofv.gencelda("","UTI BGCOLOR='#DDDDFF' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",WSPCUENTA,"si")

                  WSPIMPOTT = 0
                  WSPCANTTT = 0

              FOR WKFindex=1 TO WKFMESES

                  WSPIMPOC  = TABACCMES(CTAINDEX,WKFindex,3)
                  WSPCANTC  = TABACCMES(CTAINDEX,WKFindex,4)

                  WSPIMPOTT = WSPIMPOTT + WSPIMPOC
                  WSPCANTTT = WSPCANTTT + WSPCANTC

                  IF WSPIMPOC = 0 THEN
                     WSPIMPOC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOC = "$&nbsp;" & Formatnumber(WSPIMPOC ,2,0,0,-1) & "&nbsp;"
                  END IF



                  IF WSPCANTC = 0 THEN
                     WSPCANTC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTC = WSPCANTC & "&nbsp;"
                  END IF

                  Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","RIGHT","","","","",WSPIMPOC,"si")
                  Celda3 = Celda3 & ofv.gencelda("","uT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","RIGHT","","","","",WSPCANTC,"si")

              NEXT

                  IF WSPIMPOTT = 0 THEN
                     WSPIMPOTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOTT = "$&nbsp;" & Formatnumber(WSPIMPOTT ,2,0,0,-1) & "&nbsp;"
                  END IF



                  IF WSPCANTTT = 0 THEN
                     WSPCANTTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTTT = WSPCANTTT & "&nbsp;"
                  END IF


              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPIMPOTT,"si")
              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPCANTTT,"si")

              Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#FFFFFF' ","","","25","","",Celda3,"si")




  END IF



              WSPCUENTA = "TOTAL:"

              Celda3 = ""
              Celda3 = Celda3 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPCUENTA,"si")

                  WSPIMPOTT = 0
                  WSPCANTTT = 0

              FOR WKFindex=1 TO WKFMESES

                  WSPIMPOC  = TABACCMES(0,WKFindex,3)
                  WSPCANTC  = TABACCMES(0,WKFindex,4)

                  WSPIMPOTT = WSPIMPOTT + WSPIMPOC
                  WSPCANTTT = WSPCANTTT + WSPCANTC

                  IF WSPIMPOC = 0 THEN
                     WSPIMPOC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOC = "$&nbsp;" & Formatnumber(WSPIMPOC ,2,0,0,-1) & "&nbsp;"
                  END IF



                  IF WSPCANTC = 0 THEN
                     WSPCANTC = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTC = WSPCANTC & "&nbsp;"
                  END IF

                  Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPIMPOC,"si")
                  Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPCANTC,"si")

              NEXT

                  IF WSPIMPOTT = 0 THEN
                     WSPIMPOTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPIMPOTT = "$&nbsp;" & Formatnumber(WSPIMPOTT ,2,0,0,-1) & "&nbsp;"
                  END IF



                  IF WSPCANTTT = 0 THEN
                     WSPCANTTT = "&nbsp;-" & "&nbsp;"
                  ELSE
                     WSPCANTTT = WSPCANTTT & "&nbsp;"
                  END IF


              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPIMPOTT,"si")
              Celda3 = Celda3 & ofv.gencelda("","VT  style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","","","","",WSPCANTTT,"si")


              Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F7EEDD' ","","","25","","",Celda3,"si")


  END IF



  else



  end if


    sHTML = sHTML & "<BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa BGCOLOR='#FFFFFF' STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","95% ","","0","","0","0","","",Fila1,"si") & "</CENTER>"


'''" & XspSShw("#dddddd","125") & ";
'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "



  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sAccion = "ORMAPAEVENTOS.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER
  Celda = "" 'ofv.BotonVolver(sAccion)


  sHTML = sHTML &  ofv.botoneraESSX("0","",cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if



Response.Write sHTML



%>