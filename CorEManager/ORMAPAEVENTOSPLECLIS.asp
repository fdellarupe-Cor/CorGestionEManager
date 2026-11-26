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

    IF sORCARACTER <> "TODOS" THEN
       sORCARACTERW = " (A.TRCCARACTER = 0" & sORCARACTER & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & sORCARACTERW
       SCONECTOR = " AND "

    END IF

    IF sORRIESGO <> "TODOS" THEN
       sORRIESGOW   = " (A.TRCEVENTORIESGO1 = 0" & sORRIESGO & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & sORRIESGOW
       SCONECTOR = " AND "

    END IF



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




    IF SESSION(WSPSESION) <> "" THEN
       SESSION(WSPSESION) =  SWHERE & SESSION(WSPSESION)
    END IF

    WSPWHERE = SESSION(WSPSESION)


  ELSE

    WSPWHERE = SESSION(WSPSESION)


  END IF





  if sMODAL = "REG" or sMODAL = "ENR"  or sMODAL = "ENCA"  or sMODAL = "CSN"  then

     Strsql = "     select A.*, B.ACTIVIDAD FROM ORTRACER A "
     Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
	 
	IF (SFECHASELECCION = 1) THEN
    	Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCCARACTER, A.FECORIG , A.TRCESTADO desc , A.FECREGIS DESC  "
	ELSEIF (SFECHASELECCION = 2) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCCARACTER, A.FECCONTA , A.TRCESTADO desc , A.FECREGIS DESC  "
	ELSEIF (SFECHASELECCION = 3) THEN
		'Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCCARACTER, A.FECREGIS , A.TRCESTADO desc , A.FECREGIS DESC  "
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCCARACTER, A.TRCESTADO desc , A.FECREGIS DESC  "
	ELSEIF (SFECHASELECCION = 4) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCCARACTER, A.FECIDENT , A.TRCESTADO desc , A.FECREGIS DESC  "
	ELSEIF (SFECHASELECCION = 5) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCCARACTER, A.FECFINAL , A.TRCESTADO desc , A.FECREGIS DESC  "
	ELSEIF (SFECHASELECCION = 6) THEN
		Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCCARACTER, A.FECFINENR , A.TRCESTADO desc , A.FECREGIS DESC  "
	END IF
     
     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

     if sMODAL = "CSN"  then      WKACCION = " - LISTA DE EVENTOS RAIZ"

  'response.write strsql & "<br>"

   IF WWSORFECINIDDSC <> "" OR WWSORFECINIHDSC <> "" THEN

      Celda = ""
      Celda = Celda & ofv.GenCelda("","ut   ","","100%","","","",WWSORFECINIDDSC & WWSORFECINIHDSC,"si")
      Fila  = ofv.GenRow("","","","","","","",Celda,"si")

      sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;BORDER-COLOR:#993333;' valign='middle'  ","","100% ","20","0","","-1","1","","",Fila,"si")
      sHTML = sHTML & "<br>"

   END IF

    Fila = ""
    Fila1 = ""

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=4 BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN"),"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","7%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","53%","","","","DESCRIPCION","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","20%","","","","MONTO","si")
    
	IF (SFECHASELECCION = 1) THEN
    	Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","20%","","","","F.INICIO","si")
	ELSEIF (SFECHASELECCION = 2) THEN
		Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","20%","","","","F.REG.CONT.","si")
	ELSEIF (SFECHASELECCION = 3) THEN
		Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","20%","","","","F.CARGA","si")
	ELSEIF (SFECHASELECCION = 4) THEN
		Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","20%","","","","F.DESCUB.","si")
	ELSEIF (SFECHASELECCION = 5) THEN
		Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","20%","","","","F.FINALIZADO","si")
	ELSEIF (SFECHASELECCION = 6) THEN
		Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","20%","","","","F.EST.FIN","si")
	END IF
	
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


            WSPIMPOFTT = 0
            WSPCANTFTT = 0

        wkmontonetoT =  0
        wkmontoRECUT =  0
        wkmontoBRUTT =  0
        wkmontoOMONT =  0
        wkmontoPERDT =  0
        wkmontoDESAT =  0


  'Encabezados

   Formulario = ""
   FILAX = ""
   Celda = ""

  'response.write CHR(13) & CHR(13)
  'response.write strsql & "<br>"
  'response.write CHR(13) & CHR(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


  ELSE
  posI = 0

  xTrColor = "#f0f0f0"

  CARACTANT = "NO"
  CMESANT   = "NO"
  CANIOANT  = "NO"

  posi = 0
  Do While (Not rs.EOF)



     wkmontonetox =  rs(12) & ""

     if wkmontonetox =  "" OR wkmontonetox =  "0"  then


     else

        wkCARACTER =  CSTR(RS("TRCCARACTER") & "")
		IF (SFECHASELECCION = 1) THEN
			WSPFECHA   =  CDATE( CLNG(RS("FECORIG")) + CLNG(CDATE("01-01-1900")) )
		ELSEIF (SFECHASELECCION = 2) THEN
			WSPFECHA   =  CDATE( CLNG(RS("FECCONTA")) + CLNG(CDATE("01-01-1900")) )
		ELSEIF (SFECHASELECCION = 3) THEN
			WSPFECHA   =  CDATE( CLNG(RS("FECREGIS")) + CLNG(CDATE("01-01-1900")) )
		ELSEIF (SFECHASELECCION = 4) THEN
			WSPFECHA   =  CDATE( CLNG(RS("FECIDENT")) + CLNG(CDATE("01-01-1900")) )
		ELSEIF (SFECHASELECCION = 5) THEN
			WSPFECHA   =  CDATE( CLNG(RS("FECFINAL")) + CLNG(CDATE("01-01-1900")) )
		ELSEIF (SFECHASELECCION = 6) THEN
			WSPFECHA   =  CDATE( CLNG(RS("FECFINENR")) + CLNG(CDATE("01-01-1900")) )
		END IF
        

        gtcsanio   = datepart("yyyy",WSPFECHA)
        gtcsmes    = datepart("m",WSPFECHA)

        WSPMESE    =  gtcsmes
        WSPANIOE   =  gtcsanio

        WSPMESEX    =  CSTR(gtcsmes & "")
        WSPANIOEX   =  CSTR(gtcsanio & "")

        IF wkCARACTER <> CARACTANT THEN

           IF   CARACTANT = "NO" THEN

           ELSE

                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","uTI COLSPAN=2 BGCOLOR='#f0f0f0'    ","right","","","","","&nbsp;&nbsp;&nbsp;" & "TOTAL " & CMESANT & "/" & CANIOANT,"si")

                wspcolor = "#eeeeff"

                WSPIMPOF = wkmontoBRUTTM

				IF ORXLS="SI" THEN
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = ""
					ELSE
					   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF,"si")
				ELSE
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = "&nbsp;-"
					ELSE
					   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
	
					Celda1 = Celda1 & ofv.gencelda("","uT BGCOLOR='#99ccee' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
				END IF

                Celda1 = Celda1 & ofv.gencelda("","uT   ","","","","","","&nbsp;","si")

                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")


                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","uT COLSPAN=4   ","","","","","","&nbsp;","si")
                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")



                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#DDDDDD' COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","","TOTAL " & DSCCARACTANT,"si")

                wspcolor = "#eeeeff"

                WSPIMPOF = wkmontoBRUTTC

				IF ORXLS="SI" THEN
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = ""
					ELSE
					   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF,"si")
				ELSE
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = "&nbsp;-"
					ELSE
					   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#DDDDDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
				END IF
				
                Celda1 = Celda1 & ofv.gencelda("","uT   ","","","","","","&nbsp;","si")

                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")

           END IF


           CMESANT   = WSPMESEX
           CANIOANT  = WSPANIOEX
           wkmontoBRUTTM =  0

           CARACTANT = wkCARACTER
           DSCCARACTANT = "CARACTER " & CARACTANT
           wkmontoBRUTTC =  0

           Strsql = "Select DESCRI From ORTIPOCARACTER WHERE OBJID = 0" & wkCARACTER
           set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if NOT rsX1.eof then
              DSCCARACTANT = RSX1(0)
           END IF


           Celda1 = ""
           Celda1 = Celda1 & ofv.gencelda("","uT COLSPAN=4   ","","","","","","&nbsp;","si")
           Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")

           Celda1 = ""
           Celda1 = Celda1 & ofv.gencelda("","VTI COLSPAN=4 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",DSCCARACTANT & "<HR>","si")
           Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F7EEDD' ","","","25","","",Celda1,"si")


           Celda1 = ""
           Celda1 = Celda1 & ofv.gencelda("","uT COLSPAN=4   ","","","","","","&nbsp;","si")
           Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")

           Celda1 = ""
           Celda1 = Celda1 & ofv.gencelda("","uTI COLSPAN=4 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","","&nbsp;&nbsp;&nbsp;" & CMESANT & "/" & CANIOANT,"si")
           Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#99ccee' ","","","25","","",Celda1,"si")


        ELSE


           IF CANIOANT <> WSPANIOEX OR CMESANT <> WSPMESEX THEN


                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","uTI COLSPAN=2    ","right","","","","","&nbsp;&nbsp;&nbsp;" & "TOTAL " & CMESANT & "/" & CANIOANT,"si")

                wspcolor = "#eeeeff"

                WSPIMPOF = wkmontoBRUTTM

				IF ORXLS="SI" THEN
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = ""
					ELSE
					   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF,"si")
				ELSE
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = "&nbsp;-"
					ELSE
					   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
	
					Celda1 = Celda1 & ofv.gencelda("","uT BGCOLOR='#99ccee' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
				END IF
				
				
                Celda1 = Celda1 & ofv.gencelda("","uT   ","","","","","","&nbsp;","si")

                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")


                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","uT COLSPAN=4   ","","","","","","&nbsp;","si")
                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")


           CMESANT   = WSPMESEX
           CANIOANT  = WSPANIOEX
           wkmontoBRUTTM =  0


           Celda1 = ""
           Celda1 = Celda1 & ofv.gencelda("","uTI COLSPAN=4 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","","&nbsp;&nbsp;&nbsp;" & CMESANT & "/" & CANIOANT,"si")
           Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#99ccee' ","","","25","","",Celda1,"si")


           END IF

        END IF


        posI = 0

        xTrColor = "#f0f0f0"



        wkmontoneto =  0


        wkmontoRECU =  0
        wkmontoBRUT =  cdbl(rs(12))
        wkmontoOMON =  0
        wkmontoPERD =  0
        wkmontoDESA =  0



        wkmontonetoT =  wkmontonetoT + wkmontoneto
        wkmontoRECUT =  wkmontoRECUT + wkmontoRECU
        wkmontoBRUTT =  wkmontoBRUTT + wkmontoBRUT
        wkmontoOMONT =  wkmontoOMONT + wkmontoOMON
        wkmontoPERDT =  wkmontoPERDT + wkmontoPERD
        wkmontoDESAT =  wkmontoDESAT + wkmontoDESA

        wkmontoBRUTTC =  wkmontoBRUTTC + wkmontoBRUT
        wkmontoBRUTTM =  wkmontoBRUTTM + wkmontoBRUT

        CELDA1 = ""


        Celda1 = Celda1 & ofv.gencelda("","uTI BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",RS(0),"si")


        WKMDESCRI = RS(11) & "."

        Celda1 = Celda1 & ofv.gencelda("","UTI style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",WKMDESCRI,"si")

            wspcolor = "#eeeeff"

            WSPIMPOF = wkmontoBRUT

			IF ORXLS="SI" THEN
				IF WSPIMPOF = 0 THEN
				   WSPIMPOF = ""
				ELSE
				   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
				END IF
				Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF,"si")
			ELSE
				IF WSPIMPOF = 0 THEN
				   WSPIMPOF = "&nbsp;-"
				ELSE
				   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
				END IF
				Celda1 = Celda1 & ofv.gencelda("","uT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
			END IF


        ''    WSPFECHA = CDATE( CLNG(RS("FECORIG")) + CLNG(CDATE("01-01-1900")) )

        Celda1 = Celda1 & ofv.gencelda("","uT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPFECHA & "&nbsp;","si")

        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")
     end if



    posi = posi + 1
    rs.Movenext
  Loop


           IF   wkmontoBRUTTC > 0 THEN


                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","uTI COLSPAN=2 BGCOLOR='#f0f0f0'    ","right","","","","","&nbsp;&nbsp;&nbsp;" & "TOTAL " & CMESANT & "/" & CANIOANT,"si")

                wspcolor = "#eeeeff"

                WSPIMPOF = wkmontoBRUTTM

				IF ORXLS="SI" THEN
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = ""
					ELSE
					   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF,"si")				
				ELSE
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = "&nbsp;-"
					ELSE
					   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","uT BGCOLOR='#99ccee' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
				END IF
				
                Celda1 = Celda1 & ofv.gencelda("","uT   ","","","","","","&nbsp;","si")

                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")


                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","uT COLSPAN=4   ","","","","","","&nbsp;","si")
                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")



                Celda1 = ""
                Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#DDDDDD' COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","","TOTAL " & DSCCARACTANT,"si")

                wspcolor = "#eeeeff"

                WSPIMPOF = wkmontoBRUTTC

				IF ORXLS="SI" THEN
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = ""
					ELSE
					   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF,"si")
				ELSE
					IF WSPIMPOF = 0 THEN
					   WSPIMPOF = "&nbsp;-"
					ELSE
					   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
					END IF
					Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#DDDDDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
				END IF
				
                Celda1 = Celda1 & ofv.gencelda("","uT   ","","","","","","&nbsp;","si")

                Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")

           END IF

           Celda1 = ""
           Celda1 = Celda1 & ofv.gencelda("","VT COLSPAN=4   ","","","","","","&nbsp;","si")
           Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")





        CELDA1 = ""



        WKMDESCRI = "TOTAL:"

        Celda1 = Celda1 & ofv.gencelda("","VTI COLSPAN=2 BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",WKMDESCRI,"si")

            wspcolor = "#eeeeff"

            WSPIMPOF = wkmontoBRUTT

			IF ORXLS="SI" THEN
				IF WSPIMPOF = 0 THEN
				   WSPIMPOF = ""
				ELSE
				   WSPIMPOF = "$ " & Formatnumber(WSPIMPOF ,2,0,0,-1)
				END IF
				Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF,"si")
			ELSE
				IF WSPIMPOF = 0 THEN
				   WSPIMPOF = "&nbsp;-"
				ELSE
				   WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
				END IF
				Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")
			END IF

        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","","&nbsp;","si")

     '   Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")
    sHTML = sHTML & "<BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","80% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"
  END IF

  else
  end if
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


function TablaBusqueda()

TablaBusqueda = ""

end function
Response.Write sHTML



%>