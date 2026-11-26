<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sMODAL       = request.querystring("MODAL")
  sMODALMAIN   = request.querystring("MODALMAIN")
  sCODTRACER  = request.querystring("CODTRACER")
  sFECHA        = request.querystring("FECHA")
  sOPCION        = request.querystring("OPCION")
  sMODAL2       = request.querystring("MODAL2")
  sESTILO      = request.querystring("ESTILO") & ""
 ' RESPONSE.WRITE SMODAL
  if sMODAL = "VTR" then sMODAL = "REG"

  if sMODAL = "" then sMODAL = "REG"

     IF sESTILO <> "XTR" THEN
        sHTML = ofv.FormHeader(session("FormN"))
     ELSE
        sHTML = ofv.MENUHeader("","")
     END IF

  WKVISION = ""

  if sMODAL = "REG" then

     WKVISION = "REG"

  ELSEIF  sMODAL = "ENR" then

     WKVISION = "ENR"

  ELSEIF  sMODAL = "NO" then

     WKVISION = "REG"

  ELSEIF  sMODAL = "ACC" then

     WKVISION = "REG"

  ELSEIF  sMODAL = "IAG" then


    if sMODALMAIN = "PEN" then

       WKVISION = "AGN"

    ELSE

       WKVISION = "PEN"

    END IF

  ELSE

       WKVISION = sMODAL


  END IF


          WSPSEN2 = "<FONT COLOR='#DD0000' SIZE=3>&#0149;</FONT>"
          WSPSEN2RR = "<FONT COLOR='#000000' SIZE=3>&#0149;</FONT>"
          WSPSEN2J = "<FONT COLOR='#FF0088' SIZE=3>&#0149;</FONT>"
          WSPSEN2R = "<FONT COLOR='#880088' SIZE=3>&#0149;</FONT>"

    if sMODAL = "ENCA"  then
       WSPSESION = "ORWHERE2"
    ELSE
       WSPSESION = "ORWHERE"
    END IF



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


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

    SOREVALUADOS    = request.FORM("OREVALUADOS")
    SORRIESGOSIDEN  = request.FORM("ORRIESGOSIDEN")
    SORSPAIDEN      = request.FORM("ORSPAIDEN")
    ORESTADOS	    = ucase(""&request.FORM("ORESTADOS"))


    SORLINNEG       = request.FORM("ORLINNEG")
    SORTIPORG       = request.FORM("ORTIPORG")
    SORFALTANTES    = request.FORM("ORFALTANTES")

    SORFECINID      = request.FORM("ORFECINID")
    SORFECINIH      = request.FORM("ORFECINIH")


    '-------------------------------------------
    SORFECREGCONTABLED      = request.FORM("ORFECREGCONTABLED")
    SORFECREGCONTABLEH      = request.FORM("ORFECREGCONTABLEH")

    SORFECCARGAD      		= request.FORM("ORFECCARGAD")
    SORFECCARGAH      		= request.FORM("ORFECCARGAH")

    SORFECDESCUD      		= request.FORM("ORFECDESCUD")
    SORFECDESCUH      		= request.FORM("ORFECDESCUH")

    SORFECFIND      		= request.FORM("ORFECFIND")
    SORFECFINH      		= request.FORM("ORFECFINH")

    SORFECESTADOFIND     	= request.FORM("ORFECESTADOFIND")
    SORFECESTADOFINH      	= request.FORM("ORFECESTADOFINH")

    SORAREA					=request.FORM("ORAREA")
    '--------------------------------------------------------
    SORPERIODO      = request.FORM("ORPERIODO")

	SORCODIGOBCRA 	= request.FORM("ORCODIGOBCRA")
    'RESPONSE.WRITE SORAREA

    SORCODIGOW      = ""
    sORCARACTERW    = ""
    SORTEXTOW       = ""
    sORRIESGOW      = ""
    SORCUENTAW      = ""
    SORCLIENTEW     = ""
    SORMONTODW      = ""
    SORMONTOHW      = ""
    SORPROCESOW     = ""
    SORLOTEW        = ""
    SOREVALUADOSW   = ""
    SORRIESGOSIDENW = ""
    SORSPAIDENW     = ""

    SORLINNEGW      = ""
    SORTIPORGW      = ""
    SORFALTANTESW   = ""

    SORFECINIDW     = ""
    SORFECINIHW     = ""
    SORPERIODOw     = ""
    SORAREAW		= ""

	SORCODIGOBCRAW 	= ""

    SCONECTOR = ""
    SWHERE    = ""

    if sMODAL = "REG"  then
       SCONECTOR = " AND "
    ELSE
       SCONECTOR = " AND "
    END IF




       SESSION(WSPSESION) =  ""


    IF SORPERIODO <> "TODOS" THEN

         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & SORPERIODO
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODOMAP  = RSX1(2)
         ELSE
            WKORPERIODOMAP  = 0
         END IF

         IF SORPROCESO <> "TODOS" THEN
            SORPROCESOW   = " (A.TRCPROCESO = 0" & sORPROCESO & ") "

            SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORPROCESOW
            SCONECTOR = " AND "

         ELSE


            SORPROCESOW   = " (A.TRCPROCESO IN ("


            Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
            Strsql = Strsql & " AND codobjmaster = '0" & WKORPERIODOMAP & "' and tipobjmaster = 'PROCESO'  "
            Strsql = Strsql & " ORDER BY OBJETOS.INDOBJ "

            SCOMSEP = ""

            set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            do while not rsx1.eof
               SORPROCESOW   = SORPROCESOW  & SCOMSEP & "'" & RSX1(0) & "'"
               SCOMSEP = ","
               rsx1.movenext
            loop

            SORPROCESOW   = SORPROCESOW & " )) "

            SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORPROCESOW
            SCONECTOR = " AND "

         END IF



    END IF





    IF SORCODIGO <> "TODOS" THEN

       SORCODIGOW   = " (A.OBJID = 0" & SORCODIGO & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORCODIGOW
       SCONECTOR = " AND "

    END IF

    IF sORCARACTER <> "TODOS" THEN
       sORCARACTERW = " (A.TRCCARACTER = 0" & sORCARACTER & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & sORCARACTERW
       SCONECTOR = " AND "

    END IF


    IF SORTEXTO <> "TODOS" THEN
       SORTEXTOW    = " (A.TRCOBSERV LIKE '%" & SORTEXTO & "%') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
       SCONECTOR = " AND "

    END IF


    IF sORRIESGO <> "TODOS" THEN
       sORRIESGOW   = " (A.TRCEVENTORIESGO1 = 0" & sORRIESGO & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & sORRIESGOW
       SCONECTOR = " AND "

    END IF

    IF SORCUENTA <> "TODOS" THEN
       SORCUENTAW   = " (A.TRCCTACBLE = '" & SORCUENTA & "') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORCUENTAW
       SCONECTOR = " AND "

    END IF

    IF SORCLIENTE <> "TODOS" THEN
       SORCLIENTEW  = " (A.TRCCLIENTE = '" & SORCLIENTE & "') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORCLIENTEW
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



    IF SORLOTE <> "TODOS" THEN
       SORLOTEW   = " (A.TRCLOTEING = '" & sORLOTE & "') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORLOTEW
       SCONECTOR = " AND "

    END IF

    IF SORSPAIDEN <> "TODOS" THEN

       IF SORSPAIDEN = "1" THEN
          SORSPAIDENW   = " (A.trcprodser <> '0' AND A.trcidentact <> '0' ) "
       ELSEIF SORSPAIDEN = "2" THEN
          SORSPAIDENW   = " (A.trcprodser <> '0' AND A.trcidentact = '0' ) "
       ELSE
          SORSPAIDENW   = " (A.trcprodser = '0' AND A.trcidentact = '0' ) "
       END IF

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORSPAIDENW
       SCONECTOR = " AND "

    END IF

    IF SOREVALUADOS <> "TODOS" THEN

       IF SOREVALUADOS = "1" THEN
          SOREVALUADOSW   = " (B.ACTIVIDAD <> '') "
       ELSE
          SOREVALUADOSW   = " (B.ACTIVIDAD is null) "
       END IF

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SOREVALUADOSW
       SCONECTOR = " AND "

    END IF


    IF ORESTADOS <> "TODOS" THEN

	    SORESTADOSW   = " (A.TRCESTADO = '"& ORESTADOS &"') "

	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORESTADOSW
	       SCONECTOR = " AND "

    ELSE

	    SORESTADOSW   = " ( (A.TRCESTADO = 'NO') OR (A.TRCESTADO IS NULL) "


        Strsql = "select distinct a.codigo, a.descripcion "
        Strsql = Strsql & " from       orestados a "
        Strsql = Strsql & " inner join orestadosper b on (a.objid = b.estado) "
        Strsql = Strsql & " where a.tipo > 0 and b.permiso > 0 "

        Strsql = Strsql & " and ( "
        Strsql = Strsql & " (b.usrgrp = 'TODOS') OR "
        Strsql = Strsql & " (B.USRGRP = 'USR-" & Session("Usuario") & "')  "



        if  Session("DELLOG") = "S" then
            XMFTGRP = split(session("SegGrpLst"),chr(9))
            xcount  = ubound(XMFTGRP) + 1
            for XMFI = 0 to ubound(XMFTGRP)
                Strsql = Strsql & " OR (B.USRGRP = 'GRP-" & XMFTGRP(XMFI) & "') "
            next
        else
            Strsql = Strsql & " OR (B.USRGRP = 'GRP-" & SESSION("SEGGRP") & "') "
        end if

             Strsql = Strsql & " ) "

             set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             do while not rsx1.eof

	        SORESTADOSW   = SORESTADOSW  & " OR (A.TRCESTADO = '" & rsx1(0) & "') "
                rsx1.movenext
             loop

	       SORESTADOSW   = SORESTADOSW  & " ) "


	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORESTADOSW
	       SCONECTOR = " AND "



    END IF

    IF SORRIESGOSIDEN <> "TODOS" THEN

       IF SORRIESGOSIDEN = "1" THEN
          SORRIESGOSIDENW   = " (A.TRCEVENTO <> '0') "
       ELSE
          SORRIESGOSIDENW   = " (A.TRCEVENTO = '0') "
       END IF


       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORRIESGOSIDENW
       SCONECTOR = " AND "

    END IF


    IF SORLINNEG <> "TODOS" THEN

       IF SORLINNEG = "0" THEN

         SORLINNEGW   = " ((C.NIVEL1 = '0') OR (C.NIVEL1 IS NULL) ) "

       ELSE

         SORLINNEGW   = " (C.NIVEL1 = '"& SORLINNEG &"') "

       END IF

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORLINNEGW
       SCONECTOR = " AND "

    END IF

    IF SORTIPORG <> "TODOS" THEN

       IF SORTIPORG = "0" THEN

         SORTIPORGW   = " ((A.TRCEVENTORIESGO1 = '0') OR (A.TRCEVENTORIESGO1 IS NULL) ) "

       ELSE

         SORTIPORGW   = " (A.TRCEVENTORIESGO1 = '"& SORTIPORG &"') "

       END IF

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORTIPORGW
       SCONECTOR = " AND "

    END IF


       WWSORFECINIDDSC = ""
       WWSORFECINIHDSC = ""

    IF SORFECINID = "" AND Len(trim(SORFECINID & "")) =0  THEN SORFECINID = "TODOS"
    IF SORFECINIH = "" AND Len(trim(SORFECINIH & "")) =0  THEN SORFECINIH = "TODOS"


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


'------------------------------------------------------------------------------------------------------------------------------


    'SORFECREGCONTABLED      = request.FORM("ORFECREGCONTABLED")
   ' SORFECREGCONTABLEH      = request.FORM("ORFECREGCONTABLEH")




IF SORFECREGCONTABLED = "" AND Len(trim(SORFECREGCONTABLED & "")) =0  THEN SORFECREGCONTABLED = "TODOS"
IF SORFECREGCONTABLEH = "" AND Len(trim(SORFECREGCONTABLEH & "")) =0  THEN SORFECREGCONTABLEH = "TODOS"


    IF SORFECREGCONTABLED <> "TODOS" THEN

       if IsDate(SORFECREGCONTABLED) then
          WWSORFECINID = CLNG(CDATE(SORFECREGCONTABLED) - CLNG(CDATE("01-01-1900")))
          WWSORFECINIDDSC = " DESDE " & CDATE(SORFECREGCONTABLED)
       else
          WWSORFECINID = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
          WWSORFECINIDDSC = " DESDE " & DATE()
       end if



       SORFECINIDW   = " (NOT A.fecconta < " & WWSORFECINID & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIDW
       SCONECTOR = " AND "

    END IF

    IF SORFECREGCONTABLEH <> "TODOS" THEN

       if IsDate(SORFECREGCONTABLEH) then
          WWSORFECINIH = CLNG(CDATE(SORFECREGCONTABLEH) - CLNG(CDATE("01-01-1900")))
          WWSORFECINIHDSC = " HASTA " & CDATE(SORFECREGCONTABLEH)
       else
          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
          WWSORFECINIHDSC = " HASTA " & DATE()
       end if



       SORFECINIHW   = " (NOT A.fecconta > " & WWSORFECINIH & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
       SCONECTOR = " AND "

    END IF







    IF SORFECCARGAD = "" AND Len(trim(SORFECCARGAD & "")) =0  THEN SORFECCARGAD = "TODOS"
	IF SORFECCARGAH = "" AND Len(trim(SORFECCARGAH & "")) =0  THEN SORFECCARGAH = "TODOS"


	    IF SORFECCARGAD <> "TODOS" THEN

	       if IsDate(SORFECCARGAD) then
	          WWSORFECINID = CLNG(CDATE(SORFECCARGAD) - CLNG(CDATE("01-01-1900")))
	          WWSORFECINIDDSC = " DESDE " & CDATE(SORFECCARGAD)
	       else
	          WWSORFECINID = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
	          WWSORFECINIDDSC = " DESDE " & DATE()
	       end if



	       SORFECINIDW   = " (NOT A.fecregis < " & WWSORFECINID & ") "

	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIDW
	       SCONECTOR = " AND "

	    END IF

	    IF SORFECCARGAH <> "TODOS" THEN

	       if IsDate(SORFECCARGAH) then
	          WWSORFECINIH = CLNG(CDATE(SORFECCARGAH) - CLNG(CDATE("01-01-1900")))
	          WWSORFECINIHDSC = " HASTA " & CDATE(SORFECCARGAH)
	       else
	          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
	          WWSORFECINIHDSC = " HASTA " & DATE()
	       end if



	       SORFECINIHW   = " (NOT A.fecregis > " & WWSORFECINIH & ") "

	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
	       SCONECTOR = " AND "

    END IF






    IF SORFECDESCUD = "" AND Len(trim(SORFECDESCUD & "")) =0  THEN SORFECDESCUD = "TODOS"
	IF SORFECDESCUH = "" AND Len(trim(SORFECDESCUH & "")) =0  THEN SORFECDESCUH = "TODOS"


	    IF SORFECDESCUD <> "TODOS" THEN

	       if IsDate(SORFECDESCUD) then
	          WWSORFECINID = CLNG(CDATE(SORFECDESCUD) - CLNG(CDATE("01-01-1900")))
	          WWSORFECINIDDSC = " DESDE " & CDATE(SORFECDESCUD)
	       else
	          WWSORFECINID = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
	          WWSORFECINIDDSC = " DESDE " & DATE()
	       end if



	       SORFECINIDW   = " (NOT A.fecident < " & WWSORFECINID & ") "

	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIDW
	       SCONECTOR = " AND "

	    END IF

	    IF SORFECDESCUH <> "TODOS" THEN

	       if IsDate(SORFECDESCUH) then
	          WWSORFECINIH = CLNG(CDATE(SORFECDESCUH) - CLNG(CDATE("01-01-1900")))
	          WWSORFECINIHDSC = " HASTA " & CDATE(SORFECDESCUH)
	       else
	          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
	          WWSORFECINIHDSC = " HASTA " & DATE()
	       end if



	       SORFECINIHW   = " (NOT A.fecident > " & WWSORFECINIH & ") "

	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
	       SCONECTOR = " AND "

    END IF






    IF SORFECFIND = "" AND Len(trim(SORFECFIND & "")) =0  THEN SORFECFIND = "TODOS"
	IF SORFECFINH = "" AND Len(trim(SORFECFINH & "")) =0  THEN SORFECFINH = "TODOS"


	    IF SORFECFIND <> "TODOS" THEN

	       if IsDate(SORFECFIND) then
	          WWSORFECINID = CLNG(CDATE(SORFECFIND) - CLNG(CDATE("01-01-1900")))
	          WWSORFECINIDDSC = " DESDE " & CDATE(SORFECFIND)
	       else
	          WWSORFECINID = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
	          WWSORFECINIDDSC = " DESDE " & DATE()
	       end if



	       SORFECINIDW   = " (NOT A.fecfinal < " & WWSORFECINID & ") "

	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIDW
	       SCONECTOR = " AND "

	    END IF

	    IF SORFECFINH <> "TODOS" THEN

	       if IsDate(SORFECFINH) then
	          WWSORFECINIH = CLNG(CDATE(SORFECFINH) - CLNG(CDATE("01-01-1900")))
	          WWSORFECINIHDSC = " HASTA " & CDATE(SORFECFINH)
	       else
	          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
	          WWSORFECINIHDSC = " HASTA " & DATE()
	       end if



	       SORFECINIHW   = " (NOT A.fecfinal > " & WWSORFECINIH & ") "

	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
	       SCONECTOR = " AND "

    END IF



	    'SORFECESTADOFIND      = request.FORM("ORFECESTADOFIND")
	    'SORFECESTADOFINH      = request.FORM("ORFECESTADOFINH")



	    IF SORFECESTADOFIND = "" AND Len(trim(SORFECESTADOFIND & "")) =0  THEN SORFECESTADOFIND = "TODOS"
		IF SORFECESTADOFINH = "" AND Len(trim(SORFECESTADOFINH & "")) =0  THEN SORFECESTADOFINH = "TODOS"


		    IF SORFECESTADOFIND <> "TODOS" THEN

		       if IsDate(SORFECESTADOFIND) then
		          WWSORFECINID = CLNG(CDATE(SORFECESTADOFIND) - CLNG(CDATE("01-01-1900")))
		          WWSORFECINIDDSC = " DESDE " & CDATE(SORFECESTADOFIND)
		       else
		          WWSORFECINID = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
		          WWSORFECINIDDSC = " DESDE " & DATE()
		       end if



		       SORFECINIDW   = " (NOT A.FECFINENR < " & WWSORFECINID & ") "

		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIDW
		       SCONECTOR = " AND "

		    END IF

		    IF SORFECESTADOFINH <> "TODOS" THEN

		       if IsDate(SORFECESTADOFINH) then
		          WWSORFECINIH = CLNG(CDATE(SORFECESTADOFINH) - CLNG(CDATE("01-01-1900")))
		          WWSORFECINIHDSC = " HASTA " & CDATE(SORFECESTADOFINH)
		       else
		          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
		          WWSORFECINIHDSC = " HASTA " & DATE()
		       end if



		       SORFECINIHW   = " (NOT A.FECFINENR > " & WWSORFECINIH & ") "

		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
		       SCONECTOR = " AND "

    END IF


 IF SORAREA <> "TODOS" THEN

       SORCODIGOW   = " (A.TRCECORESP = 0" & SORAREA & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORCODIGOW
       SCONECTOR = " AND "

    END IF



'------------------------------------------------------------------------------------------------------------------------------


    IF SORFALTANTES <> "TODOS" THEN

       SORFALTANTESTD = RIGHT(SORFALTANTES,1)
       SORFALTANTESDD = LEFT(SORFALTANTES,LEN(SORFALTANTES) - 2)

       SORFALTANTESW   = " ((A." & SORFALTANTESDD & " = '0') OR (A." & SORFALTANTESDD & " IS NULL)  "

       IF SORFALTANTESTD = "X" THEN

          SORFALTANTESW   = SORFALTANTESW & "  OR (A." & SORFALTANTESDD & " = '')  "

       END IF

       SORFALTANTESW   = SORFALTANTESW & " ) "



       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFALTANTESW
       SCONECTOR = " AND "

    END IF

'***************************************************************************************************************************

	IF SORCODIGOBCRA <> "TODOS" THEN

		IF (ISNUMERIC(SORCODIGOBCRA)) THEN

			SORCODIGOBCRAW = " A.OBJID IN (SELECT OBCRA.OBJIDORTRACER FROM ORREGINFOBCRADET OBCRA WHERE OBCRA.EVENTORAIZ = " & SORCODIGOBCRA & " ) "

			SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORCODIGOBCRAW

			SCONECTOR = " AND "

		END IF

	END IF



'***************************************************************************************************************************


    if sMODAL = "ENCA"  then
       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & " A.CODTRACER <> '" & SCODTRACER & "' AND (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND TRCESTADO <> 'CAN' "
    END IF




    IF SESSION(WSPSESION) <> "" THEN
       SESSION(WSPSESION) =  SWHERE & SESSION(WSPSESION)
    END IF

    WSPWHERE = SESSION(WSPSESION)


  ELSE

    WSPWHERE = SESSION(WSPSESION)


  END IF



'RESPONSE.WRITE smodal & "<BR><BR>"


  if sMODAL = "NO" then



       ' strsql = "DELETE FROM ORTRACER WHERE TRCESTADO = 'NO' "
       sFECREGIS = CLNG(DATE()) - clng(cdate("01-01-1900")) - 1
       Strsql = "DELETE FROM ORTRACER WHERE TRCESTADO = 'NO' and TRCMARCAACTP <> 'RM' AND fecregis < " & sFECREGIS

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


     SMODAL = WKVISION

  ELSEif sMODAL = "ACC" then

        strsql = "UPDATE ORTRACER SET TRCESTADO = 'REG' WHERE CODTRACER = '" & SCODTRACER & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


     SMODAL = WKVISION
  END IF



  if sMODAL = "REG" or sMODAL = "ENR"  or sMODAL = "ENCA"  or sMODAL = "CSN"  then


  if sMODAL = "REG"  then

  ''   Strsql = "     select A.*, B.ACTIVIDAD "
  ''   Strsql = Strsql & " FROM ORTRACER A "
  ''   Strsql = Strsql & " "
  ''   Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
  ''   Strsql = Strsql & " WHERE A.TRCUSRREG = '" & SESSION("USUARIO") & "' " & WSPWHERE & " ORDER BY   A.FECREGIS DESC   "

     Strsql = "     select A.*, B.ACTIVIDAD, "

     Strsql = Strsql & "  CASE A.TRCESTADO "
     Strsql = Strsql & "        WHEN 'CAN' THEN '4' "
     Strsql = Strsql & "        WHEN 'REG' THEN '1' "
     Strsql = Strsql & "        WHEN 'ENR' THEN '2' "
     Strsql = Strsql & "        WHEN 'ENRF' THEN '2' "
     Strsql = Strsql & "        WHEN 'FIN' THEN '3' "

     Strsql = Strsql & "         END AS ESTORD "
	 'Strsql = Strsql & " CASE A.TRCESTADO WHEN 'NO' THEN 'RESGUARDO RECTIFICATIVA' "
	 'Strsql = Strsql & " ELSE ( SELECT DESCRIPCION FROM ORESTADOS WHERE CODIGO = A.TRCESTADO ) "
	 'Strsql = Strsql & " END AS ESTORD"

	 'Strsql = Strsql & " , C.NIVEL1 , A.TRCEVENTORIESGO1, D.DESCRIPCION AS ESTADODSC "
	 Strsql = Strsql & " , C.NIVEL1 , A.TRCEVENTORIESGO1"

	 Strsql = Strsql & " ,CASE A.TRCESTADO WHEN 'NO' THEN 'RESGUARDO RECTIFICATIVA' "
	 Strsql = Strsql & " ELSE ( SELECT DESCRIPCION FROM ORESTADOS WHERE CODIGO = A.TRCESTADO ) "
	 Strsql = Strsql & " END AS ESTADODSC "

     Strsql = Strsql & " "
     Strsql = Strsql & " FROM ORTRACER A "
     Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
     Strsql = Strsql & " LEFT JOIN ORLINNEGO2  C ON (A.TRCNEGOLINE = C.OBJID)  "
     Strsql = Strsql & " LEFT JOIN ORESTADOS   D ON (A.TRCESTADO = D.CODIGO)  "
     Strsql = Strsql & " WHERE A.TRCESTADO <> 'DOBLIGA' AND (A.TRCUSRREG = '" & SESSION("USUARIO") & "' "

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   if session("APLMDCLI")="CFA" THEN
     if  Session("DELLOG") = "S" then
         XMFTGRP = split(session("SegGrpLst"),chr(9))
         xcount  = ubound(XMFTGRP) + 1
         for XMFI = 0 to ubound(XMFTGRP)
             Strsql = Strsql & " OR (A.TRCGRUPOWNER = '" & XMFTGRP(XMFI) & "') "
         next
     else
         Strsql = Strsql & " OR (A.TRCGRUPOWNER = '" & SESSION("SEGGRP") & "') "
     end if
   END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     Strsql = Strsql & ") "
     Strsql = Strsql & WSPWHERE
     Strsql = Strsql & " ORDER BY  a.objid DESC  "

   ''''  Strsql = Strsql & " ORDER BY  ESTORD, A.FECREGIS DESC  "



     WKACCION = " - REGISTRACION DE EVENTOS DE PERDIDA - USUARIO: " & SESSION("SEGUSUARIO")

  ELSE

     Strsql = "     select A.*, B.ACTIVIDAD, "

     Strsql = Strsql & "  CASE A.TRCESTADO "
     Strsql = Strsql & "        WHEN 'CAN' THEN '4' "
     Strsql = Strsql & "        WHEN 'REG' THEN '1' "
     Strsql = Strsql & "        WHEN 'ENR' THEN '2' "
     Strsql = Strsql & "        WHEN 'ENRF' THEN '2' "
     Strsql = Strsql & "        WHEN 'FIN' THEN '3' "

     Strsql = Strsql & "         END AS ESTORD "
	 'Strsql = Strsql & " CASE A.TRCESTADO WHEN 'NO' THEN 'RESGUARDO RECTIFICATIVA' "
	 'Strsql = Strsql & " ELSE ( SELECT DESCRIPCION FROM ORESTADOS WHERE CODIGO = A.TRCESTADO ) "
	 'Strsql = Strsql & " END AS ESTORD"

	 'Strsql = Strsql & " , C.NIVEL1 , A.TRCEVENTORIESGO1, D.DESCRIPCION AS ESTADODSC "
	 Strsql = Strsql & " , C.NIVEL1 , A.TRCEVENTORIESGO1"

	 Strsql = Strsql & " ,CASE A.TRCESTADO WHEN 'NO' THEN 'RESGUARDO RECTIFICATIVA' "
	 Strsql = Strsql & " ELSE ( SELECT DESCRIPCION FROM ORESTADOS WHERE CODIGO = A.TRCESTADO ) "
	 Strsql = Strsql & " END AS ESTADODSC "

     Strsql = Strsql & " "
     Strsql = Strsql & " FROM ORTRACER A "
     Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
     Strsql = Strsql & " LEFT JOIN ORLINNEGO2  C ON (A.TRCNEGOLINE = C.OBJID)  "
     Strsql = Strsql & " LEFT JOIN ORESTADOS   D ON (A.TRCESTADO = D.CODIGO)  "
     Strsql = Strsql & " WHERE A.TRCESTADO <> 'DOBLIGA'  " & WSPWHERE
     Strsql = Strsql & " ORDER BY  a.objid DESC  "



     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

     if sMODAL = "ENCA"  then      WKACCION = " - ENCADENAMIENTO DE EVENTOS DE PERDIDA"
     if sMODAL = "CSN"  then      WKACCION = " - CONSULTA DE EVENTOS DE PERDIDA"

  END IF

	'response.write chr(13) & chr(13)
	'response.write strsql & "<br>"
	'response.write chr(13) & chr(13)



   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE EVENTOS DE RIESGO " & WKACCION,"si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

   Celda = ""
   Celda = Celda & ofv.GenCelda("","uti  ","","10%","","","","SELECCION: ","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","90%","","","",WSPWHERE,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")

 ' sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;BORDER-COLOR:#993333;' valign='middle'  ","","100% ","30","0","","-1","1","","",Fila,"si")
  sHTML = sHTML & "<br>"




  if sMODAL = "ENCA"  then

    sForm = "formACT"
    Formulario = ""
    sInput = "&nbsp;"
    Celda = ofv.gencelda("","VTI colspan=5 ","","","","","","EVENTO DE PERDIDA ACTUAL","si")
    Formulario = ofv.genrow("","AAA bgcolor='#F7EEDD' ","","","25","","",Celda,"si")

    Celda = ""

    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","5%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","CUENTA","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","30%","","","","CARACTER","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","15%","","","","F.REG.CONTABLE","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","MONTO","si")

    Formulario = Formulario & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ""
       Strsql2 = "select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "

''  response.write strsql & "<br>"
       set rs2 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
       if rs2.eof then

          WKCODIGO = "-----"
          WKDESCRI = "-- SIN DEFINIR --"
          WKCARAC = "-- SIN DEFINIR --"
          WKFCONT = "-- SIN DEFINIR --"
          WKMONTO = "-- SIN DEFINIR --"

       ELSE

          WKCODIGO = RS2(0)

          if (rs2(29) & "") = "" then
             dFechan = "-------"
          else
             dFechan = RS2(29)
          end if

          WKDESCRI = dFechan

          sInputx = "-------"

          strsql3 = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs2(24)
          set rsniv = ofv.crearconsultaEx(StrSql3,cn,1,parametros)
          if not rsniv.eof then sInputx = rsniv(0)
          call ofv.cerrarconsulta(rsniv)

          WKCARAC = sInputx

          if clng(rs2(5)) = 0 then
             dFechan = "-------"
          else
             dFechan = GTFechadorSH(rs2(5))
          end if
          WKFCONT = dFechan

          if cdbl(rs2(12)) = 0 then
             dFechan = "-------"
          else
             dFechan = RS2(12)
          end if
          WKMONTO = dFechan

       END IF

       Celda = Celda & ofv.gencelda("","VTI onclick='document." & sForm & ".submit();' title='ver evento.' STYLE='cursor:pointer;color:#000000;background-color:#ddddff;border-style:outset;border-width:2;' ","","","","","",WKCODIGO,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKDESCRI,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKCARAC,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKFCONT,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKMONTO,"si")

       sAccion = "ORBASERIESGOSADD.ASP?MODAL=CSN&CODTRACER=" & RS2(1)

       FILAX = ofv.genrow("","","","","25","","",Celda,"si")
       Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"_BLANK","post","",FilaX,"si")

  sHTML = sHTML & "" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#993333;' ","","98% ","","0","","0","0","","",Formulario,"si")

  sHTML = sHTML & "<BR>"

 END IF







  'Encabezados

   Formulario = ""
   FILAX = ""
   Celda = ""


 ' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     Celda = ""

     Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","","NO EXISTEN EVENTOS DE RIESGO QUE CUMPLAN CON LA SELECCION ACTUAL","si")
     Fila  = ofv.GenRow("","","","","","","",Celda,"si")

     sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")


  ELSE
  posI = 0

  xTrColor = "#f0f0f0"
 ' response.write strsql & "<br>"

  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0


  Do While (Not rs.EOF)


    sForm = "form" & posi


    Filax  = ""

    Celda = ""

    'CODIGO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' ","","5%","","","","CODIGO","si")


    'INCIDENTE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' ","","30%","","","","FECHA","si")

    'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","30%","","","","CARACTER","si")

    'CUENTA
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","30%","","","","CUENTA","si")

  ''  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""


    'CODIGO
    SINPUT = RS(0)


 IF sESTILO <> "XTR" THEN

  if sMODAL = "REG"  then


    sAccion = "ORBASERIESGOSADD.ASP?MODAL=VTR&CODTRACER=" & RS(1) & "&CantRegMover=" & CantRegMoverX & ""

  ELSEif sMODAL = "ENCA"  OR sMODAL = "CSN" then


    sAccion = "ORBASERIESGOSADD.ASP?MODAL=CSN&CODTRACER=" & RS(1)

  ELSE


    sAccion = "ORBASERIESGOSADD.ASP?MODAL=ENR&CODTRACER=" & RS(1) & "&CantRegMover=" & CantRegMoverX & ""

  END IF

    Celda = Celda & ofv.gencelda("","vT onclick='document." & sForm & ".submit();' title='ver evento.' style='cursor:pointer;background-color:#ddddff;border-style:outset;border-width:2;' ROWSPAN=3  ","center","","","","",sInput,"si")

 else

    Celda = Celda & ofv.gencelda("","vT  style='background-color:#ddddff;border-style:outset;border-width:2;' ROWSPAN=3  ","center","","","","",sInput,"si")

 end if



    'F.INCIDENTE
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","FECHA:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","25%","","","",sInput,"si")



    'CARACTER

    sInputx = "-------"

    strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs(24)
    set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsniv.eof then sInputx = rsniv(0)
    call ofv.cerrarconsulta(rsniv)

    sInput = sInputx
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CARACTER:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



    'CUENTA CONTABLE
    if (rs(29) & "") = "" then
      SINPUT = "-------"
    else
      SINPUT = RS(29)
    end if
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CUENTA:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


  if sMODAL = "ENCA"  then
     SINPUT = "<A CLASS=CC HREF='ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=ENR&OPCION=SI&PRETRACER=" & RS(1) & "&MODAL2="&SMODAL2&"' >RELACIONAR</A>"
     Celda = Celda & ofv.gencelda("","vT  title='Relacionar este evento.' style='cursor:pointer;background-color:#ddddff;border-style:outset;border-width:2;' ROWSPAN=3  ","center","","","","",sInput,"si")
  END IF

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""
    'CLIENTE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' ","","","","","","CLIENTE","si")

    'MONTO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","MONTO","si")

    'ESTADO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","ESTADO","si")

  '  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""
    'CLIENTE
    if (rs(28) & "") = "" then
      SINPUT = "-------"
    else
      SINPUT = RS(28)
    end if
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CLIENTE:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'MONTO
   ' response.write rs(16).NAME
    if not isnumeric(REPLACE(rs(16)&"",",",".")) then
     SINPUT = "-------"
    else

    if cdbl( rs(16)) = 0 then
      SINPUT = "-------"
    else
      SINPUT = RS(16)
    end if
    end if
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","MONTO NETO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'ESTADO
    wkestado = rs(17) & ""
    wkmodoOT = rs("TRCMARCAACTP") & ""

    wkestadoD = rs("ESTADODSC") & ""

    IF wkestado = "NO" AND wkmodoOT = "RM" THEN

        wkestadoD = "<font color='#000000' >RESGUARDO RECTIFICATIVA</font>"

    ELSEIF (wkestadoD & "") = "" THEN

        wkestadoD = "<font color='#000099' >EN REVISION</font>"

    ELSE

        wkestadoD = "<font color='#009900' >" & wkestadoD & "</font>"

    END IF



    sInput = wkestadoD

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","ESTADO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","VTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""
    'DESCRIPCION
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' COLSPAN=3 ","","","","","","DESCRIPCION","si")

  '  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""
    'OBSERVACIONES

    wwwdescri = RS(11) & ""

 ''''   if len(wwwdescri) > 60 then wwwdescri = left(wwwdescri,60) & "...(+)"

    if (rs(43) & "") = "S"    then wwwdescri = WSPSEN2 & "&nbsp;" & wwwdescri & "&nbsp;"
    if (rs(43) & "") = "RM"   then wwwdescri = WSPSEN2RM & "&nbsp;" & wwwdescri & "&nbsp;"
    if (rs(43) & "") = "J"    then wwwdescri = WSPSEN2J & "&nbsp;" & wwwdescri & "&nbsp;"
    if (rs(43) & "") = "R" OR (rs(43) & "") = "XR"  OR (rs(43) & "") = "KR"  then wwwdescri = WSPSEN2R & "&nbsp;" & wwwdescri & "&nbsp;"


    SINPUT = wwwdescri
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","DESCRIPCION:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Formulario = Formulario & Filax

 IF sESTILO <> "XTR" THEN


    sInput = "&nbsp;"
    Celda = ofv.gencelda("","UT colspan=7 style='border-TOP-style:SOLID;border-TOP-width:2;border-TOP-color:#000000;' ","","","","","",sInput,"si")

    Filax  = ofv.genrow("","AAA bgcolor='#fff9f9' ","","","25","","",Celda,"si")

 else

    Filax  = ""


 end if


 '''''   Formulario = Formulario & Filax

  if sMODAL = "ENCA" OR sMODAL = "CSN"  then

      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"_BLANK","post","",FilaX,"si")

  ELSE

      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")

  END IF



    posi = posi + 1
    rs.Movenext

    IF sESTILO <> "XTR" THEN

       IF NOT (posi < CantRegMostrarx) THEN EXIT DO

    END IF

  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar AND sMODAL = "REG" AND sESTILO <> "XTR" then
    sAccion = "ORBASERIESGOSADD.asp?MODAL=AAG"
    Agregar = ofv.BotonAgregar(sAccion)
  end if
  Buscar  = ""
  Celda   = ""
  IF sMODAL2="VTR" THEN SMODAL=SMODAL2
  sAccion = "orbaseriesgos.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER
  Celda = ofv.BotonVolver(sAccion)

  WSPREEMP = "modal=" & smodal  & "&CODTRACER=" & SCODTRACER & "&"

    IF sESTILO <> "XTR" THEN

       sHTML = sHTML &  replace(ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda),"cantregmostrar=",WSPREEMP & "cantregmostrar=")

    ELSE

       sHTML = sHTML &  ofv.botoneraESSX("0","","","","no","no",posi,"","","","","","")

   END IF

  call ofv.cerrarconsulta(rs)







  else



  end if


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>