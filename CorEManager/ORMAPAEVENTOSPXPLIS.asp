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

  if sMODAL = "" then sMODAL = "CSN"

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
          TABRIESGOS(1,10) = "NO ES PRODUCTO" 'DESCRI
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



  Strsql = "Select * From ORTIPOPRODUCTO  ORDER BY DESCRI"
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

    SORPRODUCTO  = request.FORM("ORPRODUCTO")
    SORTEVENTO   = request.FORM("ORTEVENTO")

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

    SORPRODUCTOW  = ""
    SORTEVENTOW   = ""

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

    IF SORPRODUCTO <> "TODOS" THEN
       SORPRODUCTOW   = " (A.TRCPRODUCTO = '0" & SORPRODUCTO & "') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORPRODUCTOW
       SCONECTOR = " AND "

    END IF

    IF SORTEVENTO <> "TODOS" THEN
       SORTEVENTOW   = " (A.TRCEVENTORIESGO1 = '0" & SORTEVENTO & "') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORTEVENTOW
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
	
	
    'SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '')  AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECORIG > 0 "
	SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '')  AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' "




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
     Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCESTADO desc , A.FECREGIS DESC  "
     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

     if sMODAL = "CSN"  then      WKACCION = " - PERDIDAS POR PRODUCTO"


	'response.write(chr(13) & chr(13))
	'response.write strsql
	'response.write(chr(13) & chr(13))




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


  'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


  ELSE
  posI = 0

  xTrColor = "#f0f0f0"



  posi = 0
  Do While (Not rs.EOF)



     wkmontonetox =  rs(16) & ""

     if wkmontonetox =  "" then


     else



        wkmontoneto =  cdbl(rs(16))

        wktipoevento = rs(30)
        'wkPRODUCTO   = rs(25)
		wkCARACTER	= RS("TRCCARACTER")


        wkmontonetoP =  cdbl(wkmontoneto)

        ITABRIESGOS1 = 10
        ITABRIESGOS2 = 10


        for ii=11 to TTABRIESGOS1

            if cstr( wktipoevento & "") = cstr(TABRIESGOS(ii,0) & "") then
               TABRIESGOS(ii,2) = TABRIESGOS(ii,2) + 1
               ITABRIESGOS1 = ii
               exit for
            end if

        next

        for ii=11 to TTABRIESGOS2

            if cstr(wkCARACTER & "") = cstr(TABRIESGOS(0,ii) & "") then
               TABRIESGOS(2,ii) = TABRIESGOS(2,ii) + 1
               ITABRIESGOS2 = ii
               exit for
            end if

        next
		
		

        if  ITABRIESGOS1 = 10 then TABRIESGOS(10,2) = TABRIESGOS(10,2) + 1

        if  ITABRIESGOS2 = 10 then TABRIESGOS(2,10) = TABRIESGOS(2,10) + 1



        TABRIESGOS(ITABRIESGOS1,ITABRIESGOS2) = TABRIESGOS(ITABRIESGOS1,ITABRIESGOS2) + wkmontonetoP

        TABRIESGOS(ITABRIESGOS1,4) = TABRIESGOS(ITABRIESGOS1,4) + wkmontonetoP

        TABRIESGOS(4,ITABRIESGOS2) = TABRIESGOS(4,ITABRIESGOS2) + wkmontonetoP


     end if





    posi = posi + 1
    rs.Movenext
  Loop
  END IF


  else



  end if

''''''''''''''''''''' VERSION 1

   TTABIMPA = TTABRIESGOS2
   TTABIMPAX = TTABRIESGOS2 - 9
   ITABIMPA = 11

   if TABRIESGOS(2,10) > 0 then

      TTABIMPAX = TTABRIESGOS2 - 8
      ITABIMPA = 10

   end if


   TTABPROB = TTABRIESGOS1
   ITABPROB = 11

   if TABRIESGOS(10,2) > 0 then

      ITABPROB = 10

   end if

''' encabezados

    Fila = ""
    Fila1 = ""


    TTABIMPAX = 1


    MASTTOTAL = 0


    FOR IIMPA=ITABIMPA TO TTABIMPA

            MASTTOTAL = MASTTOTAL + TABRIESGOS(4,IIMPA)
    NEXT



    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'   ","center","","","","","TIPO DE EVENTO","si")
    FOR II=ITABIMPA TO TTABIMPA


        if TABRIESGOS(4,II) = 0 then
            WKDISPLAY = "none"
        else
            TTABIMPAX = TTABIMPAX + 1
            Celda = Celda & ofv.gencelda("","VTI BGCOLOR='#336699' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",TABRIESGOS(1,II),"si")
        end if


    NEXT
    Filan = ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg  COLSPAN=" & TTABIMPAX + 1 & " BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN"),"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")



    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' COLSPAN=" & TTABIMPAX & " ","CENTER","","","","","PRODUCTO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' VALIGN='BOTTOM' ROWSPAN=2 ","CENTER","10%","","","","&nbsp;TOTAL&nbsp;","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si") & Filan



            WSPIMPOFTT = 0
            WSPCANTFTT = 0

    FOR IPROB=ITABPROB TO TTABPROB



        CELDA1 = ""


        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'  ","","20%","","","",TABRIESGOS(IPROB,1),"si")



        FOR IIMPA=ITABIMPA TO TTABIMPA

            wspcolor = "#f0f0f0"

            WSPIMPOF = TABRIESGOS(IPROB,IIMPA)
            if MASTTOTAL=0 then MASTTOTAL=1

            WSPPORC  = WSPIMPOF / MASTTOTAL * 100
            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "&nbsp;%"


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               wspcolor = "#ddddff"
             '''  WSPIMPOF = "<A CLASS=NT style='COLOR:#000000;' HREF='' TARGET='_BLANK' >$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "</A>"

               'WSPIMPOF = "<A CLASS=TT style='COLOR:#000000;' HREF='javascript:void(0);' TARGET='_BLANK' >$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "&nbsp;</A>"
			   WSPIMPOF = "<A CLASS=TT style='COLOR:#000000;' HREF='ORMAPAEVENTOSLISDET.ASP?ORLINNEG=" & TABRIESGOS(0,IIMPA) & "&ORTIPORG=" & TABRIESGOS(IPROB,0) & "' TARGET='_BLANK' ><U>$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "&nbsp;</U></A>"
			RESPONSE.Write(CHR(13) & CHR(13))
			RESPONSE.Write(TABRIESGOS(0,IIMPA) & CHR(13) & IIMPA)
			RESPONSE.Write(CHR(13) & CHR(13))
			RESPONSE.Write(TABRIESGOS(IPROB,0) & CHR(13) & IPROB)
			RESPONSE.Write(CHR(13) & CHR(13))
               WSPIMPOF = WSPIMPOF & "<HR SIZE=1 STYLE='border-style:solid;border-width:1;border-color:#993333;' >" & WSPPORC & "&nbsp;"

            END IF

			Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")

            if TABRIESGOS(4,IIMPA) = 0 then
               WKDISPLAY = "none"
            else
               Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")
            end if




        NEXT

            wspcolor = "#eeeeff"

            WSPIMPOF = TABRIESGOS(IPROB,4)
             if MASTTOTAL=0 then MASTTOTAL=1
            WSPPORC  = WSPIMPOF / MASTTOTAL * 100
            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "&nbsp;%"


            WSPIMPOFTT = WSPIMPOFTT + WSPIMPOF


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "<b>$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "</b>"
               WSPIMPOF = WSPIMPOF & "<HR SIZE=1 STYLE='border-style:solid;border-width:1;border-color:#0000aa;' >" & WSPPORC & "&nbsp;"
			   
			'RESPONSE.Write(CHR(13) & CHR(13))
			'RESPONSE.Write(WSPIMPOF & CHR(13))
			'RESPONSE.Write(CHR(13) & CHR(13))
            END IF




        Celda1 = Celda1 & ofv.gencelda("","uT BGCOLOR='#F0F0F0' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")


    NEXT


        CELDA1 = ""



        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'   ","RIGHT","","","","","TOTAL&nbsp;","si")




        FOR IIMPA=ITABIMPA TO TTABIMPA

            wspcolor = "#eeeeff"





            WSPIMPOF = TABRIESGOS(4,IIMPA)
            WSPCANTF = 0

            WSPPORC  = WSPIMPOF / MASTTOTAL * 100
            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "&nbsp;%"


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "<b>$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "</b>"
               WSPIMPOF = WSPIMPOF & "<HR SIZE=1 STYLE='border-style:solid;border-width:1;border-color:#0000aa;' >" & WSPPORC & "&nbsp;"
            END IF


            if TABRIESGOS(4,IIMPA) = 0 then
               WKDISPLAY = "none"
            else

              Celda1 = Celda1 & ofv.gencelda("","uT style=';COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")
            end if



        NEXT


            IF WSPIMPOFTT = 0 THEN
               WSPIMPOFTT = "&nbsp;-"
            ELSE
               WSPIMPOFTT = "$&nbsp;" & Formatnumber(WSPIMPOFTT ,2,0,0,-1)
            END IF




            Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOFTT & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")






  ''  sHTML = sHTML & "<BR>"

  ''  sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","90% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"


'''''''''''''''''''''''''''''''

''''''''''''''''''''' VERSION 2

   TTABIMPA = TTABRIESGOS2
   TTABIMPAX = TTABRIESGOS2 - 9
   ITABIMPA = 11

   if TABRIESGOS(2,10) > 0 then

      TTABIMPAX = TTABRIESGOS2 - 8
      ITABIMPA = 10

   end if


   TTABPROB = TTABRIESGOS1
   ITABPROB = 11

   if TABRIESGOS(10,2) > 0 then

      ITABPROB = 10

   end if

''' encabezados

    Fila = ""
    Fila1 = ""


    TTABIMPAX = 1


    MASTTOTAL = 0


    FOR IIMPA=ITABPROB TO TTABPROB

            MASTTOTAL = MASTTOTAL + TABRIESGOS(IIMPA,4)
    NEXT



    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'   ","center","","","","","PRODUCTOS","si")



  '''  FOR II=ITABIMPA TO TTABIMPA

    FOR II=ITABPROB TO TTABPROB

        TTABIMPAX = TTABIMPAX + 1
        Celda = Celda & ofv.gencelda("","VTI BGCOLOR='#336699' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",TABRIESGOS(II,1),"si")

    NEXT
	
	'RESPONSE.Write("PRIMER FOR " & TTABIMPAX)
	
    Filan = ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg  COLSPAN=" & TTABIMPAX + 1 & " BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN"),"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")



    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' COLSPAN=" & TTABIMPAX & " ","CENTER","","","","","TIPOS DE EVENTO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' VALIGN='BOTTOM' ROWSPAN=2 ","CENTER","10%","","","","&nbsp;TOTAL&nbsp;","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si") & Filan



            WSPIMPOFTT = 0
            WSPCANTFTT = 0

  '''  FOR IPROB=ITABPROB TO TTABPROB

	'response.write("Que vale esto " & ITABIMPA)

    FOR IPROB=ITABIMPA TO TTABIMPA

        CELDA1 = ""


     if TABRIESGOS(4,IPROB) > 0 then


    ' ELSE


        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'  ","","20%","","","",TABRIESGOS(1,IPROB),"si")



     '''   FOR IIMPA=ITABIMPA TO TTABIMPA

        FOR IIMPA=ITABPROB TO TTABPROB

            wspcolor = "#f0f0f0"

            WSPIMPOF = TABRIESGOS(IIMPA,IPROB)

            WSPPORC  = WSPIMPOF / MASTTOTAL * 100
            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "&nbsp;%"


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE
               wspcolor = "#ddddff"
			WSPIMPOF = "<A CLASS=TT style='COLOR:#000000;' HREF='ORMAPAEVENTOSPXPLISDET.ASP?ORLINNEG=" &  TABRIESGOS(0,(IPROB)) & "&ORTIPORG=" & TABRIESGOS((IIMPA),0) & "' TARGET='_BLANK' ><U>$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "&nbsp;</U></A>"
			'RESPONSE.Write(CHR(13) & CHR(13))
			'RESPONSE.Write(TABRIESGOS(0,(IIMPA)) & CHR(13))
			'RESPONSE.Write(TABRIESGOS(0,(IPROB)) & CHR(13))
			'RESPONSE.Write(WSPIMPOF & CHR(13))
			'RESPONSE.Write(CHR(13) & CHR(13))
               'WSPIMPOF = "<B>$&nbsp;" & Formatnumber(WSPIMPOF ,2,-1,0,-1) & "&nbsp;</B>"
               WSPIMPOF = WSPIMPOF & "<HR SIZE=1 STYLE='border-style:solid;border-width:1;border-color:#993333;' >" & WSPPORC & "&nbsp;"

            END IF

            Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")





        NEXT

            wspcolor = "#eeeeff"

            WSPIMPOF = TABRIESGOS(4,IPROB)

            WSPPORC  = WSPIMPOF / MASTTOTAL * 100
            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "&nbsp;%"


            WSPIMPOFTT = WSPIMPOFTT + WSPIMPOF


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "<b>$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "</b>"
               WSPIMPOF = WSPIMPOF & "<HR SIZE=1 STYLE='border-style:solid;border-width:1;border-color:#0000aa;' >" & WSPPORC & "&nbsp;"
            END IF




        Celda1 = Celda1 & ofv.gencelda("","uT BGCOLOR='#F0F0F0' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")

      END IF


    NEXT


        CELDA1 = ""

        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'   ","RIGHT","","","","","TOTAL&nbsp;","si")

        FOR IIMPA=ITABPROB TO TTABPROB

            wspcolor = "#eeeeff"

            WSPIMPOF = TABRIESGOS(IIMPA,4)
            WSPCANTF = 0
             if MASTTOTAL=0 then MASTTOTAL=1
            WSPPORC  = WSPIMPOF / MASTTOTAL * 100
            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "&nbsp;%"


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "<b>$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1) & "</b>"
               WSPIMPOF = WSPIMPOF & "<HR SIZE=1 STYLE='border-style:solid;border-width:1;border-color:#0000aa;' >" & WSPPORC & "&nbsp;"
            END IF


            Celda1 = Celda1 & ofv.gencelda("","uT style=';COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")

        NEXT


            IF WSPIMPOFTT = 0 THEN
               WSPIMPOFTT = "&nbsp;-"
            ELSE
               WSPIMPOFTT = "$&nbsp;" & Formatnumber(WSPIMPOFTT ,2,0,0,-1)
            END IF




            Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOFTT & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")






    sHTML = sHTML & "<BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","90% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"











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