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

  DIM TABMUESTRA(100,5)
  TTABMUESTRA = 0
  TTABMUESTRATOPE = 5

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

    SORMILES    = request.FORM("ORMILES")
    SORRANKING  = request.FORM("ORRANKING")

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
    SORRANKINGW         = ""

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


    IF SORRANKING <> "TODOS" THEN
       TTABMUESTRATOPE = CINT(SORRANKING)
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

	
	
    'SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' AND A.FECORIG > 0 "
	SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO' "




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
     Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCIMPNETO DESC  "
     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

     if sMODAL = "CSN"  then      WKACCION = " - PRINCIPALES PERDIDAS"


	'response.write chr(13) & chr(13)
	'response.write strsql 
	'response.write chr(13) & chr(13) & chr(13) & chr(13)





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

   WSPCANTTT = 0
   WSPCANTTT2 = 0

''' encabezados

    Fila = ""
    Fila1 = ""

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=4 BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN"),"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


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

''  DIM TABDISTRIB(100,5)

        ITABDISTRIB = 1
        TTABDISTRIB = 1
        WKDISTRIBCERO = 0
        WKDISTRIBTOTAL = 0

        wkmontoneto =  cdbl(rs(16))

        wktipoevento = rs(30)
        wkdistrib    = rs(37)
        wklineaN2    = rs(8)

        WKFEVENTODSC = "SIN DETERMINAR"

        for ii=11 to TTABRIESGOS1

            if cstr( wktipoevento & "") = cstr(TABRIESGOS(ii,0) & "") then
         ''      TABRIESGOS(ii,2) = TABRIESGOS(ii,2) + 1
               WKFEVENTODSC = TABRIESGOS(ii,1)
               ITABRIESGOS1 = ii
               exit for
            end if

        next

        WSPIMPOFTT = WSPIMPOFTT + wkmontoneto
        WSPCANTTT = WSPCANTTT + 1



        IF TTABMUESTRA < TTABMUESTRATOPE THEN

           WSPIMPOFTT2 = WSPIMPOFTT2 + wkmontoneto
           WSPCANTTT2 = WSPCANTTT2 + 1

           TTABMUESTRA = TTABMUESTRA + 1

           TABMUESTRA(TTABMUESTRA,1) = RS(11) & "&nbsp;"
           TABMUESTRA(TTABMUESTRA,2) = WKFEVENTODSC
           TABMUESTRA(TTABMUESTRA,3) = wkmontoneto
           TABMUESTRA(TTABMUESTRA,4) = 0


         END IF



        Strsql = "Select NIVEL1 From ORLINNEGO2 WHERE OBJID = 0" & wklineaN2
		
		'response.write chr(13) & chr(13)
		'response.write strsql 
		'response.write chr(13) & chr(13)
	
        set rsX3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsX3.eof then
           wklinea1    = rsX3(0)
        ELSE
           wklinea1    = 0
        END IF


        TABDISTRIB(ITABDISTRIB,1) = wklinea1
        TABDISTRIB(ITABDISTRIB,2) = CDBL(100)

        IF (wkdistrib & "") = "S" THEN
           TABDISTRIB(ITABDISTRIB,2) = CDBL("0" & RS(38))
           WKDISTRIBTOTAL = WKDISTRIBTOTAL +  TABDISTRIB(ITABDISTRIB,2)

           IF   TABDISTRIB(ITABDISTRIB,2) = 0 THEN
                 WKDISTRIBCERO =  WKDISTRIBCERO + 1

           END IF


           Strsql = "  select * FROM ORDISTRIBLN WHERE CODTRACER = '" & RS(1) & "' ORDER BY PORCENTAJE DESC"
		   
			'response.write chr(13) & chr(13)
			'response.write strsql 
			'response.write chr(13) & chr(13)

'''  response.write strsql & "<br>"
           set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if NOT rs3.eof then

              DO WHILE NOT RS3.EOF

                 ITABDISTRIB = ITABDISTRIB + 1
                 TTABDISTRIB = ITABDISTRIB

                 wklineaN2D    = RS3(1)

                 Strsql = "Select NIVEL1 From ORLINNEGO2 WHERE OBJID = 0" & wklineaN2D
				 'response.write chr(13) & chr(13)
				'response.write strsql 
				'response.write chr(13) & chr(13)
                 set rsX3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if not rsX3.eof then
                    wklinea1D    = rsX3(0)
                 ELSE
                    wklinea1D    = 0
                 END IF

                 TABDISTRIB(ITABDISTRIB,1) = wklinea1D
                 TABDISTRIB(ITABDISTRIB,2) = CDBL(RS3(2))

                 WKDISTRIBTOTAL = WKDISTRIBTOTAL +  TABDISTRIB(ITABDISTRIB,2)

                 IF   TABDISTRIB(ITABDISTRIB,2) = 0 THEN
                      WKDISTRIBCERO =  WKDISTRIBCERO + 1
                 END IF

                RS3.MOVENEXT
              LOOP
           END IF


           WKADISTRIB = TTABDISTRIB - WKDISTRIBCERO

           IF WKDISTRIBCERO > 0 THEN

              IF WKADISTRIB = 0 THEN
                 WKDISTRIBTOTAL = CDBL(100 / TTABDISTRIB )

                 FOR YYY=1 TO TTABDISTRIB
                     TABDISTRIB(YYY,2) = WKDISTRIBTOTAL
                 NEXT



              ELSE

                 WKDISTRIBTOTAL = CDBL(100 - WKDISTRIBTOTAL )
                 WKDISTRIBTOTAL = CDBL(WKDISTRIBTOTAL / WKDISTRIBCERO )

                 FOR YYY=1 TO TTABDISTRIB
                     IF TABDISTRIB(YYY,2) = 0 THEN
                        TABDISTRIB(YYY,2) = WKDISTRIBTOTAL
                     END IF
                 NEXT


              END IF


           END IF



        END IF


       FOR YYY=1 TO TTABDISTRIB

        wkmontonetoP =  cdbl(wkmontoneto * TABDISTRIB(YYY,2) / 100)

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

            if cstr(TABDISTRIB(YYY,1) & "") = cstr(TABRIESGOS(0,ii) & "") then
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


     '''   response.write yyy & "-" & wkmontoneto & "-" & TABDISTRIB(YYY,2) & "-" &  wkmontonetoP & "<br>" & chr(13)



      NEXT


     end if





    posi = posi + 1
    rs.Movenext
  Loop


  FOR II=1 TO  TTABMUESTRA


        CELDA1 = ""

        Celda1 = Celda1 & ofv.gencelda("","UTI style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","40%","","","",TABMUESTRA(II,1),"si")

        Celda1 = Celda1 & ofv.gencelda("","UTI BGCOLOR='#CCCCCC' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","40%","","","",TABMUESTRA(II,2),"si")

            wspcolor = "#eeeeff"

            WSPIMPOF = TABMUESTRA(II,3)
            WSPIMPOFP = TABMUESTRA(II,3)

            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               IF SORMILES = "1000" THEN WSPIMPOF = WSPIMPOF / 1000
               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF


            WSPPORC1 = WSPIMPOFP / WSPIMPOFTT * 100

            IF WSPPORC1 = 0 THEN
               WSPPORC1 = "&nbsp;-"
            ELSE

               WSPPORC1 = Formatnumber(WSPPORC1 ,2,-1,0,-1)

            END IF


        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","10%","","","",WSPIMPOF & "&nbsp;","si")
        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","10%","","","",WSPPORC1 & "%","si")


        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")


    NEXT 


       CELDA1 = ""

        WSPDSC = "TOTAL&nbsp;" & WSPCANTTT2 & "&nbsp;EVENTOS&nbsp;PRINCIPALES&nbsp;"

        IF SORMILES = "1000" THEN WSPDSC = WSPDSC & "(EN MILES)"

        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' COLSPAN=2 style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'   ","RIGHT","","","","",WSPDSC,"si")


            WSPIMPOFTT2P   = WSPIMPOFTT2
            WSPIMPOFTTP    = WSPIMPOFTT

            IF WSPIMPOFTT2 = 0 THEN
               WSPIMPOFTT2 = "&nbsp;-"
            ELSE
               IF SORMILES = "1000" THEN WSPIMPOFTT2 = WSPIMPOFTT2 / 1000
               WSPIMPOFTT2 = "$&nbsp;" & Formatnumber(WSPIMPOFTT2 ,2,0,0,-1)
            END IF

            WSPPORC2 = WSPIMPOFTT2P / WSPIMPOFTTP * 100

            IF WSPPORC2 = 0 THEN
               WSPPORC2 = "&nbsp;-"
            ELSE

               WSPPORC2 = Formatnumber(WSPPORC2 ,2,-1,0,-1)

            END IF


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOFTT2 & "&nbsp;","si")
        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPPORC2 & "%","si")

        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")


       CELDA1 = ""

        WSPDSC = "TOTAL&nbsp;" & (WSPCANTTT - WSPCANTTT2) & "&nbsp;RESTANTES&nbsp;"

        IF SORMILES = "1000" THEN WSPDSC = WSPDSC & "(EN MILES)"

        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' COLSPAN=2 style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'   ","RIGHT","","","","",WSPDSC,"si")

            WSPIMPOFTT3    = WSPIMPOFTTP - WSPIMPOFTT2P
            WSPIMPOFTT3P   = WSPIMPOFTT3


            IF WSPIMPOFTT3 = 0 THEN
               WSPIMPOFTT3 = "&nbsp;-"
            ELSE
               IF SORMILES = "1000" THEN WSPIMPOFTT3 = WSPIMPOFTT3 / 1000
               WSPIMPOFTT3 = "$&nbsp;" & Formatnumber(WSPIMPOFTT3 ,2,0,0,-1)
            END IF

            WSPPORC3 = WSPIMPOFTT3P / WSPIMPOFTTP * 100


            IF WSPPORC3 = 0 THEN
               WSPPORC3 = "&nbsp;-"
            ELSE

               WSPPORC3 = Formatnumber(WSPPORC3 ,2,-1,0,-1)

            END IF


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOFTT3 & "&nbsp;","si")
        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPPORC3 & "%","si")

        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")



       CELDA1 = ""

        WSPDSC = "TOTAL&nbsp;GENERAL&nbsp;"

        IF SORMILES = "1000" THEN WSPDSC = WSPDSC & "(EN MILES)"

        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' COLSPAN=2 style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'   ","RIGHT","","","","",WSPDSC,"si")

            IF WSPIMPOFTT = 0 THEN
               WSPIMPOFTT = "&nbsp;-"
            ELSE
               IF SORMILES = "1000" THEN WSPIMPOFTT = WSPIMPOFTT / 1000
               WSPIMPOFTT = "$&nbsp;" & Formatnumber(WSPIMPOFTT ,2,0,0,-1)
            END IF


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#CCDDF7' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOFTT & "&nbsp;","si")
        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#CCDDF7' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","","100%","si")

        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")


    sHTML = sHTML & "<BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","60% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"


  END IF
















  else



  end if



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

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=2 BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN"),"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")











    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'   ","center","50%","","","","TIPO DE EVENTO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","50%","","","","&nbsp;MONTO&nbsp;","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


            WSPIMPOFTT = 0
            WSPCANTFTT = 0

    FOR IPROB=ITABPROB TO TTABPROB



        CELDA1 = ""

'''ORMAPAEVENTOSPPDETLIS.asp




        Celda1 = Celda1 & ofv.gencelda("","UTI BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","20%","","","",TABRIESGOS(IPROB,1),"si")





            wspcolor = "#eeeeff"

            WSPIMPOF = TABRIESGOS(IPROB,4)


            WSPIMPOFTT = WSPIMPOFTT + WSPIMPOF


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               IF SORMILES = "1000" THEN WSPIMPOF = WSPIMPOF / 1000

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

               WSPIMPOF = "<A CLASS=NT HREF='ORMAPAEVENTOSPPDETLIS.asp?TEVENTO=" & TABRIESGOS(IPROB,0) & "' TITLE='Detalle de eventos para " & TABRIESGOS(IPROB,1) & ".' TARGET='_BLANK' STYLE='COLOR:#000000;' ><U>" & WSPIMPOF & "</U></A>"

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")


    NEXT


        CELDA1 = ""

        WSPDSC = "TOTAL&nbsp;"

        IF SORMILES = "1000" THEN WSPDSC = WSPDSC & "(EN MILES)"


        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'   ","RIGHT","","","","",WSPDSC,"si")


            IF WSPIMPOFTT = 0 THEN
               WSPIMPOFTT = "&nbsp;-"
            ELSE
            '   WSPIMPOFTT = "$&nbsp;" & WSPIMPOFTT

               IF SORMILES = "1000" THEN WSPIMPOFTT = WSPIMPOFTT / 1000

               WSPIMPOFTT = "$&nbsp;" & Formatnumber(WSPIMPOFTT ,2,0,0,-1)

            END IF




            Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOFTT & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")






  '  sHTML = sHTML & "<BR>"

  '  sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","60% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"


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