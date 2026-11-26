<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

  'response.write ("Smenu: "& Request.querystring("menu") & "<br>")
  'response.write ("DESDE: "&Request.querystring("Desde") & "<br>")
  'response.write ("HASTA: "&Request.querystring("Hasta") & "<br>")
  'response.write ("fecha: "&Request.querystring("fecha") & "<br>")
  'response.write ("SID : " & request.querystring("Id")& "<br>")
  'response.write ("SID2 : " & request.querystring("Id2")& "<br>")
  'response.write ("nacta : " & request.querystring("nacta")& "<br>")

  'response.write request.querystring









if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else




     Generar = ""
     Addqry = ""
     Response.expires = 0
     Server.scripttimeout = 600
     Response.buffer = True


     XXXX    = NEWEMPRE("NO")
     smenu   = Request.querystring("Menu")
     sorden  = Request.querystring("sorden")
     sNombre = request.querystring("Nombre")



  if Session("CorApli") = "COREM20" then
     wkaplic = "EM"
  elseif Session("CorApli") = "COREP20" then
     wkaplic = "PL"
  else
     wkaplic = "OR"
  end if


     If Len(sorden) = 0 Then
         sorden = "N"
     End If

     Sfiltro = Request.form("sfiltro")

        If Len(Sfiltro) = 0 Then
           Sfiltro = Request.querystring("sfiltro")
        End If

     If Len(Sfiltro) = 0 Then
         Sfiltro = "TODOS"
     End If

     Sfiltro = Replace(Sfiltro, "_", " ")
     CantRegMostrar = 10
     CantRegMover = CDbl(Request.querystring("CantRegMover"))
     ReDim Preserve bconsultar(4)
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "4"
     bconsultar(4) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "2"
     bconsultar(2) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "1"
     bconsultar(1) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "3"
     bconsultar(3) = OFV.AConsultar

     If smenu = "PORFECHA" OR smenu = "PORREVISION" Then
        sverestado = " and csporfechas.objnivel >= " & Session("SegNivel") & " And Estcod In (9,"
     Else
        sverestado = " and Objetos.objnivel >= " & Session("SegNivel") & " And Estados.objid In (9,"
     End If

     For i = 1 To 4
         If bconsultar(i) Then
            sverestado = sverestado & i & ","
            if i = 1 then sverestado = sverestado & "6" & ","
            if i = 2 then sverestado = sverestado & "5" & ","

         end if
     Next
     sverestado = Left(sverestado, Len(sverestado) - 1) & ") "

''''''     FILTRO POR EMPRESA

      addfrom =  ""


	If smenu = "PORFECHA" OR smenu = "PORREVISION" or smenu = "PORACTA"  Then

		if (SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1) or (SESSION("EMPMUN") = "N") THEN
			addfrom =  " ,objetos , tipobjeto  "
			sverestado = sverestado & " and csporfechas.CODOBJETO = OBJETOS.CODOBJETO "
			sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
			sverestado = sverestado & " and OBJETOS.OBJUN = " & Session("UN") & " "
		ELSE
			addfrom =  " ,OBJMEMPALC ,objetos , tipobjeto "
			sverestado = sverestado & " and csporfechas.CODOBJETO = OBJETOS.CODOBJETO "
			sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
			sverestado = sverestado & " and "
			sverestado = sverestado & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
			sverestado = sverestado & "   ( OBJETOS.OBJUN = 1 AND "
			sverestado = sverestado & "     csporfechas.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
			sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
			sverestado = sverestado & " ) "
		END IF

		sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "
		sverestado = sverestado & " and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	else


		if (SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1) or (SESSION("EMPMUN") = "N") THEN
			sverestado = sverestado & " and Objetos.OBJUN = " & Session("UN") & " "
		ELSE
			addfrom    =  " ,OBJMEMPALC "
			sverestado = sverestado & " and "
			sverestado = sverestado & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
			sverestado = sverestado & "   ( OBJETOS.OBJUN = 1 AND "
			sverestado = sverestado & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
			sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
			sverestado = sverestado & " ) "
		END IF

		sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "
		sverestado = sverestado & " and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	end if


 ' SI TRAE LOS DATOS POR FORMULARIO O QUERYSTRING

    If Request.Form.Count > 1 Then
       Robj = "F"

    else
       Robj = "Q"
    end if


   '  ********* CONSULTA POR REVISION ASGINACION DE VARIABLES
    If smenu = "PORREVISION" Then


        cdesde    = Request.form("Desde") & ""
        chasta    = Request.form("Hasta") & ""
        if cdesde="" and chasta="" then
         cdesde    = Request.querystring("Desde") & ""
         chasta    = Request.querystring("Hasta") & ""
        end if

        Addqry = "&Desde=" & cdesde   & "&Hasta=" & chasta

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&Desde=" & cdesde & "&Hasta=" & chasta
        sMenuMas    = sMenuMas & sID2

        xsid        = Replace(sID, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID  = "") Then  sID  = "1-1-1950"

        xsid        = Replace(sID2, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"


   ' *********************************************************
    ElseIf smenu = "PORFECHA" Then


        'cdesde    = Request.querystring("Desde") & ""
        'chasta   = Request.querystring("Hasta") & ""
        'Sid3   = Request.querystring("Fecha") & ""
        'Addqry = "&Desde=" & cdesde & "&Hasta=" & chasta & "&Fecha=" & fecha

        'sMenuAdd    = ""
        'sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
        'sMenuMas    = sMenuMas & sID2 & "&Fecha=" & sID3


         sID    = Request.querystring("Desde") & ""
	 sID2   = Request.querystring("Hasta") & ""
	 sID3   = Request.querystring("Fecha") & ""

	 cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
	 chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

	 Addqry =  "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
	 sMenuAdd    = ""
	 sMenuMas    = "&menu=" & smenu & "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
	 sMenuMas    = sMenuMas & sID2 & "&Fecha=" & sID3








    ElseIf smenu = "PORPRODUCTO" Then

        'If  Robj = "F" Then

        '    sID    = Request.Form("ListaS") & ""
        '    sID2   = Request.Form("ListaG") & ""
        '    sID3   = Request.Form("ListaP") & ""
        '    Addqry = "&ListaS=" & sID & "&ListaG=" & sID2 & "&ListaP=" & sID3
        '    sID    = Replace(sID, "_", " ")
        '    sID2   = Replace(sID2, "_", " ")
        '    sID3   = Replace(sID3, "_", " ")

        'Else
            sID    = Request.querystring("ListaS") & ""
            sID2   = Request.querystring("ListaG") & ""
            sID3   = Request.querystring("ListaP") & ""
            Addqry = "&ListaS=" & sID & "&ListaG=" & sID2 & "&ListaP=" & sID3
            sID    = Replace(sID, "_", " ")
            sID2   = Replace(sID2, "_", " ")
            sID3   = Replace(sID3, "_", " ")

        'End If

    ElseIf smenu = "PORNROCIRCULAR" Then

        'If  Robj = "F" Then

        '    sID    = Request.Form("ListaG") & ""
        '    sID2   = Request.Form("ListaN") & ""
        '    Addqry = "&ListaG=" & sID & "&ListaN=" & sID2
        '    sID    = Replace(sID, "_", " ")
        '    sID2   = Replace(sID2, "_", " ")
        'Else

            sID    = Request.querystring("ListaG") & ""
            sID2   = Request.querystring("ListaN") & ""
            Addqry = "&ListaG=" & sID & "&ListaN=" & sID2
            sID    = Replace(sID, "_", " ")
            sID2   = Replace(sID2, "_", " ")

        'End If

    ElseIf smenu = "PORINSTRUMENTO" Then

        If  Robj = "F" Then

            sID    = Request.Form("Lista") & ""
            sID2   = Request.Form("buscar") & ""
            Addqry = "&Lista=" & sID & "&buscar=" & sID2
            sID2   = Replace(sID2, "_", " ")

        Else

            sID    = Request.querystring("Lista") & ""
            sID2   = Request.querystring("buscar") & ""
            Addqry = "&Lista=" & sID & "&buscar=" & sID2
            sID2   = Replace(sID2, "_", " ")

        End If


    ElseIf smenu = "PORTEXTO" Then


        If  Robj = "F" Then


            sID    = Replace(Request.Form("buscar") & "", "'", "")
            sID    = Replace(sID, Chr(34), "")
            sID2   = Request.Form("lista") & ""
            tsel   = Request.Form("tsel") & ""
            tgrp   = Request.Form("tgrp") & ""
            tcod   = Request.Form("tcod") & ""
            ttip   = Request.Form("ttip") & ""
            tnam   = Request.Form("tnam") & ""

            tdoc   = Request.Form("tdoc") & ""


            Addqry = "&buscar=" & replace(sID," ","_")

            Addqry = Addqry & "&Lista=" & sID2

            Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

            Addqry = Addqry & "&tdoc=" & tdoc

            sID    = Replace(sID, "_", " ")

        Else

            sID    = Request.querystring("buscar") & ""

            sID2   = Request.querystring("Lista") & ""

            tsel   = Request.querystring("tsel") & ""
            tgrp   = Request.querystring("tgrp") & ""
            tcod   = Request.querystring("tcod") & ""
            ttip   = Request.querystring("ttip") & ""
            tnam   = Request.querystring("tnam") & ""

            tdoc   = Request.querystring("tdoc") & ""

            Addqry = "&buscar=" & sID

            Addqry = Addqry & "&Lista=" & sID2

            Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

            Addqry = Addqry & "&tdoc=" & tdoc

            sID    = Replace(sID, "_", " ")

        End If




        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&buscar=" & sID & "&Lista="
        sMenuMas = sMenuMas & sID2 & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam
        sMenuMas = sMenuMas & "&tdoc=" & tdoc


    ElseIf smenu = "PORCATEGO" Then


        If  Robj = "F" Then


            sID    = Replace(Request.Form("buscar") & "", "'", "")
            sID    = Replace(sID, Chr(34), "")

            sID2   = Request.Form("lista") & ""

            tsel   = Request.Form("tsel") & ""
            tgrp   = Request.Form("tgrp") & ""
            tcod   = Request.Form("tcod") & ""
            ttip   = Request.Form("ttip") & ""
            tnam   = Request.Form("tnam") & ""

            tdoc   = Request.Form("tdoc") & ""


            Addqry = "&buscar=" & replace(sID," ","_")

            Addqry = Addqry & "&Lista=" & sID2

            Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

            Addqry = Addqry & "&tdoc=" & tdoc

            sID    = Replace(sID, "_", " ")

        Else

            sID    = Request.querystring("buscar") & ""

            sID2   = Request.querystring("Lista") & ""

            tsel   = Request.querystring("tsel") & ""
            tgrp   = Request.querystring("tgrp") & ""
            tcod   = Request.querystring("tcod") & ""
            ttip   = Request.querystring("ttip") & ""
            tnam   = Request.querystring("tnam") & ""

            tdoc   = Request.querystring("tdoc") & ""

            Addqry = "&buscar=" & sID

            Addqry = Addqry & "&Lista=" & sID2

            Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

            Addqry = Addqry & "&tdoc=" & tdoc

            sID    = Replace(sID, "_", " ")

        End If



       ' sMenuAdd = ""
       ' sMenuMas = "&smenu=" & smenu & "&Instrumento=" & sID & "&buscar="
       ' sMenuMas = sMenuMas & sID2

        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&buscar=" & sID & "&Lista="
        sMenuMas = sMenuMas & sID2 & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam
        sMenuMas = sMenuMas & "&tdoc=" & tdoc


    ElseIf smenu = "PORIDENTIFICACION" Then

        'If  Robj = "F" Then
	'    sID    = Request.Form("buscar") & ""
        '    Addqry = "&buscar=" & sID

        'Else
            sID    = Request.querystring("buscar") & ""
            Addqry = "&buscar=" & sID
        'End If

        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&IDENTIFICACION=" & sID



    ElseIf smenu = "PORSECTOR" Then

        If  Robj = "F" Then
         SPerfil = Request.form("perfil") & ""



		 sID     = Request.form("lista") & ""
         sID2    = Request.querystring("Lista") & ""
         sID3    = Request.Form("buscar") & ""
        else



                   SPerfil = Request.querystring("perfil") & ""



                   sID     = Request.querystring("ID") & ""
                   sID2    = Request.querystring("Lista") & ""

                   sID3    = Request.querystring("buscar") & ""
        end if
                   Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil & "&id="&Sid
                   SPerfil = Replace(SPerfil, "_", " ")


    Else

        If  Robj = "F" Then

            SPerfil = Request.Form("perfil") & ""
            sID     = Request.Form("lista") & ""
            sID2    = Request.Form("lista") & ""
            sID3    = Request.Form("buscar") & ""

            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil

            SPerfil = Replace(SPerfil, "_", " ")

        Else

            SPerfil = Request.querystring("perfil") & ""
            sID     = Request.querystring("Lista") & ""
            sID2    = Request.querystring("Lista") & ""
            sID3    = Request.querystring("buscar") & ""
            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
            SPerfil = Replace(SPerfil, "_", " ")

        End If

        sID         = OFV.convertircar(sID, "_", " ")
        If sID = "" Then sID = sID3

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
        sMenuMas    = sMenuMas & sID3 & "&perfil=" & SPerfil

    End If

    Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre
     'response.write Addqry
    starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry
    sVolver = "ListadoConsultas.asp?Menu=" & smenu & sMenuAdd
   '  response.write svolver

    Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")



    Set cn = OFV.conectar(OFV.strconn2)


    strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
    strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, "
    strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, "
    strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
    strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,"
    strsql = strsql & "Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,"
    strsql = strsql & "OBJETOS.OBJUN,OBJPRTY.fecpublic "







    Select Case smenu
        Case "PORCOMPONENTE"


	    strsql = strsql & "FROM OBJACCION, COMPLEMENTOS, OBJETOS, TIPOBJETO, OBJPRTY, Estados, objlnk, objestdet  " & addfrom
            strsql = strsql & " where   "
            strsql = strsql & "OBJETOS.codobjeto = OBJlnk.codobjeto and "
            strsql = strsql & "objlnk.ltipo = 'KL' and "
            strsql = strsql & "complementos.codcomp = objlnk.codobjetol and "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
            strsql = strsql & " and Complementos.descri = '" & sID & "'"





            sTitulo = "Componente: " & sID



   '********** ARMAR SQL CONSULTA POR ESTADO

        Case "PORESTADO"

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
            strsql = strsql & " where  "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
            strsql = strsql & "and OBJPRTY.estado = " & sID
            sTitulo = "Estado: " & sID

            sqle = "Select * from estados where objid =  " & sID

           Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")

           If Not rsx.EOF Then sTitulo = "Estado: " & rsx(1)


          CASE "PORACTA"


		sID    = Request.querystring("Desde") & ""
		sID2   = Request.querystring("Hasta") & ""
		sID3   = Request.querystring("Fecha") & ""
		nacta  = Request.querystring("nacta") & ""

		cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

		Addqry = "&menu=" & smenu & "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&nacta=" & nacta
		sMenuAdd    = ""
		sMenuMas    = "&menu=" & smenu & "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&nacta=" & nacta
		sMenuMas    = sMenuMas & sID2 & "&Fecha=" & sID3


		if nacta <> "" then
			buscaacta = " Objprty.nroacta = '" & nacta & "'"
		else
			buscaacta = "Objprty.nroacta is not null"
		end if


			 cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
			 chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

		addfrom = addfrom & ", csrelpath a, tema t, prodserv  st"


		strsql = "select DISTINCT csporfechas.*, Objprty.fecacta,objetos.descri,tipobjeto.descri,objprty.fecemis,tipobjeto.descri from csporfechas join Objprty on csporfechas.codobjeto = Objprty.codobjeto, Estados  " & addfrom
		strsql = strsql & " where " & buscaacta & " and Objprty.fecacta between '" & cdesde & "' and '" & chasta & "' "


		wherefrom = " and (a.cobj = OBJETOS.codobjeto) "
		wherefrom = wherefrom & " and (a.tema = t.codtema) "
		wherefrom = wherefrom & " and (a.stema = st.codprodserv) "


		sTitulo = "Por acta:  " & nacta &"&nbsp;&nbsp; Desde: "& Sid & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & Sid2

	        starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry
    		sVolver = "ListadoConsultas.asp?Menu=" & smenu & sMenuAdd




        Case "PORTEXTO"

            htwhere = ""

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "

            if sid2 <> "99999" then

               strsql = strsql & ", eco, OBJSECTOR "

            end if

            if tsel <> "T" and tgrp <> "M" and (tgrp = "T" and tcod <> "99999") then

               strsql = strsql & ", csrelpath rp "

               if tgrp = "T" then

                  htwhere = " and OBJETOS.codobjeto = rp.cobj and rp.tema = 0" & tcod & " "

               elseif tgrp = "S" then

                  htwhere = " and OBJETOS.codobjeto = rp.cobj and rp.stema = 0" & tcod & " "

               else

                  htwhere = " and OBJETOS.codobjeto = rp.cobj and ("
                  htwhere = htwhere & " (rp.mc1 = 0" & tcod & " and rp.mt1 = '" & ttip & "' ) or "
                  htwhere = htwhere & " (rp.mc2 = 0" & tcod & " and rp.mt2 = '" & ttip & "' ) or "
                  htwhere = htwhere & " (rp.mc3 = 0" & tcod & " and rp.mt3 = '" & ttip & "' ) or "
                  htwhere = htwhere & " (rp.mc4 = 0" & tcod & " and rp.mt4 = '" & ttip & "' ) or "
                  htwhere = htwhere & " (rp.mc5 = 0" & tcod & " and rp.mt5 = '" & ttip & "' ) or "
                  htwhere = htwhere & " (rp.mc6 = 0" & tcod & " and rp.mt6 = '" & ttip & "' ) or "
                  htwhere = htwhere & " (rp.mc7 = 0" & tcod & " and rp.mt7 = '" & ttip & "' ) or "
                  htwhere = htwhere & " (rp.cobj = 0" & tcod & " and rp.tobj = '" & ttip & "' )  "
                  htwhere = htwhere & " ) "

               end if



            end if




            strsql = strsql & addfrom
            strsql = strsql & " where   "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

          '   response.write sid2
            if sid2 <> "99999" then

               strsql = strsql & " and ECO.eco = OBJSECTOR.codsector and "
               strsql = strsql & " OBJETOS.codobjeto = OBJSECTOR.codobjeto "
               strsql = strsql & " and eco.eco = '" & sID2 & "'"

            end if

            if tdoc <> "99999" then

               strsql = strsql & " and OBJETOS.tipobjeto = '" & tdoc & "' "

            end if


            if tsel <> "T" and tgrp <> "M" then

               strsql = strsql & htwhere

            end if

            nwsid = vertexto(sid)


            strsql = strsql & " and (OBJETOS.descriabr  COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%' "
            strsql = strsql & " or OBJETOS.descri   COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%') "

            'Response.write nwsid
            sTitulo = "Texto: " & sID

            if sid2 <> "99999" then

               sqle = "Select descri from eco where eco =0" & sID2
               Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
               if not rsx.eof then
                sTitulo = sTitulo & "&nbsp;|&nbsp;Sector: " & rsx(0)
               else
                sTitulo = sTitulo & "&nbsp;|&nbsp;Sector: "
               end if
               call ofv.cerrarconsulta(rsx)
            end if

            if tsel <> "T" and tgrp <> "M" then

               sTitulo = sTitulo & "&nbsp;|&nbsp;Carpeta:&nbsp;" & replace(tnam,"_","&nbsp;")


            end if

        Case "PORCATEGO"

            htwhere = ""

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, OBJRSKREL "


'''''EVENRSK



            strsql = strsql & addfrom
            strsql = strsql & " where  "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto AND "
            strsql = strsql & "OBJETOS.codobjeto = OBJRSKREL.codobjeto AND "
            strsql = strsql & "OBJRSKREL.EVENRSK = 0" & sID2


            Sqlx = "SELEct descri FROM RSKsituacion WHERE id = 0" & tdoc
            Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
            If Not rsx.EOF Then situnn = rsx(0) & ""
            Call OFV.cerrarconsulta(rsx)

            Sqlx = "SELEct descri FROM RSKcatego WHERE id = 0" & tcod
            Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
            If Not rsx.EOF Then catenn = rsx(0) & ""
            Call OFV.cerrarconsulta(rsx)

            Sqlx = "SELEct descri FROM RSKevento WHERE id = 0" & sID2
            Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
            If Not rsx.EOF Then evennn = rsx(0) & ""
            Call OFV.cerrarconsulta(rsx)

            'Response.write strsql & "<br>"
            sTitulo = ""
            sTitulo = sTitulo & "&nbsp;Evento:&nbsp;" & evennn & "<br>"
            sTitulo = sTitulo & "&nbsp;|&nbsp;Categoria:&nbsp;" & catenn
            sTitulo = sTitulo & "&nbsp;|&nbsp;Situacion:&nbsp;" & situnn







        Case "PORPRODUCTO"

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, OBJTSEGAGR " & addfrom
            strsql = strsql & " where   "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
            strsql = strsql & "   "

                If UCase(sID3 & "") <> UCase("nsnc") Then

                   strsql = strsql & " AND OBJETOS.TSEGMENTO = " & CDbl(sID3) & " "
                End If

                If UCase(sID & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
                   strsql = strsql & " AND ((OBJETOS.codobjeto = OBJTSEGAGR.codobjeto)  AND ( "
                End If

            If UCase(sID & "") <> UCase("nsnc") Then
                strsql = strsql & " (OBJTSEGAGR.TSEGMENTO = " & CDbl(sID) & " ) "
            End If
            If UCase(sID2 & "") <> UCase("nsnc") Then
                   If UCase(sID & "") <> UCase("nsnc") Then strsql = strsql & " AND "
                   strsql = strsql & " (OBJTSEGAGR.TAGRUPA = " & CDbl(sID2) & " ) "
                End If

                If UCase(sID & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
                   strsql = strsql & " ) )"
                End If


            xinstrum = ""
            If UCase(sID & "") <> UCase("nsnc") Then
                Sqlx = "SELECT descri from tiposegmento where objid=" & CDbl(sID)
                Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
                If Not rsx.EOF Then xinstrum = rsx(0) & ""
                Call OFV.cerrarconsulta(rsx)
            End If


            xinstrum2 = ""
            If UCase(sID2 & "") <> UCase("nsnc") Then
                Sqlx = "SELECT descri from tipoAGRUPA where objid=" & CDbl(sID2)
                Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
                If Not rsx.EOF Then xinstrum2 = rsx(0) & ""
                Call OFV.cerrarconsulta(rsx)
            End If
            sTitulo = "SEGMENTO :" & xinstrum & "&nbsp;|&nbsp;GRUPO :" & xinstrum2

        Case "PORNROCIRCULAR"

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, OBJTSEGAGR " & addfrom
            strsql = strsql & " where  "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

                If UCase(sID2 & "") <> UCase("nsnc") Then
                   strsql = strsql & " AND OBJETOS.codiden = '" & CStr(sID2 & "") & "' "
                End If

                If UCase(sID & "") <> UCase("nsnc") Then
                    strsql = strsql & " AND ((OBJETOS.codobjeto = OBJTSEGAGR.codobjeto)  AND (OBJTSEGAGR.TAGRUPA = " & CDbl(sID) & " ) ) "
                End If
                  ''' Response.write strsql & "<br>"

            xinstrum2 = ""
            sTitulo = " GRUPO :" & xinstrum2

        Case "PORINSTRUMENTO"



                strsql = strsql & " FROM OBJETOS  "
                strsql = strsql & " INNER JOIN OBJPRTY on "
                strsql = strsql & " OBJETOS.codobjeto = OBJPRTY.codobjeto "
                strsql = strsql & " INNER JOIN TIPOBJETO ON "
                strsql = strsql & " OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
                strsql = strsql & " INNER JOIN Estados ON "
                strsql = strsql & " OBJPRTY.estado = Estados.objid "
                strsql = strsql & " LEFT JOIN OBJPRTYADV ON "
                strsql = strsql & " OBJPRTY.codobjeto = OBJPRTYADV.codobjeto "

           '     if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 2 THEN
                   strsql = strsql & " LEFT join OBJMEMPALC "
                   strsql = strsql & " on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO "
          '     END IF


            Strsql = strsql & " WHERE  "

            strsql = strsql & " ( "

            If sID2 <> "" Then
                strsql = strsql & "  ((OBJPRTY.instrumen=" & CDbl("0" & sID) & ")"
            Else
                strsql = strsql & "  (OBJPRTY.instrumen=" & CDbl("0" & sID) & " or OBJPRTYADV.instrua=" &  CDbl("0" & sID) & ")  "
            End If

            If sID2 <> "" Then
                strsql = strsql & " AND (OBJPRTY.instid ='" & sID2 & "')) OR "
                strsql = strsql & " (OBJPRTYADV.instrua =0" & sID & " and OBJPRTYADV.instid ='" & sID2 & "')"
            End If

            strsql = strsql & " ) "

            xinstrum = ""
            Sqlx = "SELECT descri from instrumen where objid=" &  CDbl("0" & sID)
            Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
            If Not rsx.EOF Then xinstrum = rsx(0)
            Call OFV.cerrarconsulta(rsx)
            sTitulo = "INSTRUMENTO " & xinstrum & " " & sID2

        Case "PORIDENTIFICACION"

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
            strsql = strsql & " where   "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
            strsql = strsql & "and OBJETOS.Codiden = '" & sID & "' "
            sTitulo = "IDENTIFICACION " & sID

     ' ************ ARMAR SQL CONSULTA POR SECTOR

        Case "PORSECTOR"

        'Sperfil = request.querystring ("perfil")
        'Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
        'Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre
        'starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry
    	'sVolver = "ListadoConsultas.asp?Menu=" & smenu & sMenuAdd
        'response.write "Sperfil:" & Sperfil
	'response.write "SID:" & SID
	'response.write "Addqry:" & Addqry


            Select Case SPerfil
                Case "nsnc"

                    strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, eco, OBJSECTOR " & addfrom
                    strsql = strsql & " where  "
                    strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
                    strsql = strsql & "OBJPRTY.estado = Estados.objid and "
                    strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto and "
                    strsql = strsql & "ECO.eco = OBJSECTOR.codsector and "
                    strsql = strsql & "OBJETOS.codobjeto = OBJSECTOR.codobjeto "
                    strsql = strsql & "and eco.eco = '" & sID & "'"
                Case Else
                    strsql = strsql & "FROM objcontrol,objproced,OBJETOS, TIPOBJETO, OBJPRTY, Estados, eco, OBJSECTOR " & addfrom
                    strsql = strsql & " where    "
                    strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
                    strsql = strsql & "OBJPRTY.estado = Estados.objid and "
                    strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto and "
                    strsql = strsql & "ECO.eco = OBJSECTOR.codsector and "
                    strsql = strsql & "OBJETOS.codobjeto = OBJSECTOR.codobjeto "
                    strsql = strsql & "and eco.eco = '" & sID & "'"

                    strsql = strsql & " and (( objproced.codobjeto=OBJETOS.codobjeto "
                    strsql = strsql & " and objproced.codsector=OBJSECTOR.OBJID"
                    strsql = strsql & " and objproced.codperfil='" & SPerfil & "')"
                    strsql = strsql & " or (objcontrol.codobjeto=OBJETOS.codobjeto "
                    strsql = strsql & " and objcontrol.codsector=OBJSECTOR.codsector "
                    strsql = strsql & " and objcontrol.codperfil='" & SPerfil & "'))"
            End Select

            'sTitulo = "Sector:&nbsp;" & sID
            sqle = "Select descri from eco where eco =0" & sID
		   Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
		   if not rsx.eof then
			sTitulo = "Sector:&nbsp;" & rsx(0)
		   else
			sTitulo = "Sector:&nbsp;"
		   end if
               call ofv.cerrarconsulta(rsx)

            if  SPerfil <> "nsnc" then sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & SPerfil
	'***********************************************************

        Case "PORTEMA"
            sql = "Select descri From Objetos Where codobjeto = " & sID
            Set rsObj = OFV.crearconsultaEx(sql, cn, 1, "")
            If Not rsObj.EOF Then sObj = rsObj(0)
            Call OFV.cerrarconsulta(rsObj)

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
            strsql = strsql & " where    "
            strsql = strsql & "OBJETOS.tipobjeto=TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
            strsql = strsql & "and Objetos.codobjeto = " & sID
            sTitulo = "Tema: " & sObj


        Case "PORFECHA"
                      strsql = "select *  from csporfechas " & addfrom
	             cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
	             chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
	             If sID3 = "Emision" Then
	                 strsql = strsql & " WHERE   xfecemis between '" & cdesde & "' and '" & chasta & "' "
	                 sTitulo = "Creacion&nbsp;&nbsp;Desde:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & sID2
	             ElseIf sID3 = "Vigencia" Then
	                 strsql = strsql & " WHERE   xfecvigen between '" & cdesde & "' and '" & chasta & "' "
	                 sTitulo = "Vigencia&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
	             ElseIf sID3 = "Eliminacion" Then
	                 strsql = strsql & " WHERE   xfecbaja between '" & cdesde & "' and '" & chasta & "' "
	                 sTitulo = "Eliminacion&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
	             ElseIf sID3 = "Vencimiento" Then
	                 strsql = strsql & " WHERE  xfecvto between '" & cdesde & "' and '" & chasta & "' "
	                 sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
	             ElseIf sID3 = "Publicacion" Then
	                 strsql = strsql & " WHERE   xfecpublic between '" & cdesde & "' and '" & chasta & "' "
	                 sTitulo = "Publicacion&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2


            End If
	'************  ARMAR SQL CONULTA POR REVISION
        Case "PORREVISION"
            strsql = "select *  from csporfechas " & addfrom

	    strsql = strsql & "WHERE  xfecvto between '" & cdesde & "' and '" & chasta & "' "



            sTitulo = "Revision&nbsp;&nbsp;Del :&nbsp;" & cdate(cdesde)+cdate("01-01-1900") & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & cdate(chasta)+cdate("01-01-1900")

	 '****************************************
    End Select

    '******* CONDICIONES DE  FILTROS Y ORDENAMIENTO
    strsql = strsql & sverestado & wherefrom
    sTitOrd  = ""
    sTitOrd2 = ""

    If smenu = "PORFECHA" or smenu = "PORREVISION"  Then

        If Sfiltro <> "TODOS" Then
           sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
           strsql  = strsql & " and csporfechas.descri='" & Sfiltro & "' "
        End If

        Select Case sorden
                Case "N"
                        sTitOrd = sTitOrd & "Nombre"
                        strsql  = strsql & " order by objdescri, csporfechas.descri"
                Case "T"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by  csporfechas.descri,objdescri"
                Case "C"

                        sTitOrd = sTitOrd & "Creacion"
                        strsql  = strsql & " order by  fecemis,objdescri, csporfechas.descri"

                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by  estdescri,objdescri, csporfechas.descri"
        End Select
    Else



        If Sfiltro <> "TODOS" Then
           sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
           strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
        End If

        Select Case sorden
                Case "N"
                        sTitOrd = sTitOrd & "Nombre"
                        strsql  = strsql & " order by objetos.descri,tipobjeto.descri "
                Case "T"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by tipobjeto.descri,objetos.descri "
                Case "C"
                        sTitOrd = sTitOrd & "Creacion"
                        strsql  = strsql & " order by objprty.fecemis,tipobjeto.descri "
                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by estdescri,objetos.descri,tipobjeto.descri "
        End Select
    End If


   '**********************************************************
    response.write "<br><br><br>" & strsql & "<br>" & smenu



    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

     'rsArray = rs.GetRows()
     'nr = UBound(rsArray, 2) + 1
     'response.write "There were " & nr & " matches."



 ''   sHTML = OFV.MENUHeader("","")

 '  response.write "<br><br><br><br>" & Request.form & "<br>"

    sHTML = ofv.FormHeader(session("FormN")) & chr(13)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs


 '   response.write Request.querystring & "<br>"
 '   response.write Request.form & "<br>"

  sImagenM = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
  sImagena = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>4</FONT></B>"



  Fila    = ""
  Celda   = ""
  sHref   = ""

  IF  Session("EMPMUN") = "S" THEN

  sLink   = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
  Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","","EMPRESA:&nbsp;&nbsp;","si")
  Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
  Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

  Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

  Tabla   = ofv.GenTabla("","aaa  ","","100% ","","0","","-1","-2","","",Fila,"si")
  sHTML   = sHTML & Tabla
  sHTML   = sHTML & "<br>"
  end if

  Fila = ""



 '***************** LINK DEL CAMBIAR VISTA
    sNombre=replace(sNombre," ","_")
    Select Case smenu

	Case "PORACTA"
    	   	sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& Request.querystring("Desde")  &"&Hasta="& Request.querystring("hasta") &"&Fecha="& Request.querystring("fecha")
    	   	sAccion = sAccion &"&nacta="& nacta

    	Case "PORFECHA"
		sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& Request.querystring("Desde")  &"&Hasta="& Request.querystring("hasta") &"&Fecha="& Request.querystring("fecha")

    	Case "PORESTADO"
    	   	sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")

    	Case "PORREVISION"
    		sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")

      	Case "PORSECTOR"
     		sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&perfil="& Sperfil &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")

      	Case "PORCOMPONENTE"
         	sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre

        Case "PORTEXTO"

	      	sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

        Case "PORPRODUCTO"


         sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

  	Case "PORIDENTIFICACION"


         	sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

        Case "PORNROCIRCULAR"


        sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry


    Case "PORINSTRUMENTO"


        sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

    End Select


   'response.write "<br> starget:"& starget



   '*****************************************
   '''
    'sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre& "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")
    sInput   = ofv.GenerarInput("","Vista Web","submit","","10","bt","")

   ' sInput   = sinput & ofv.GenerarInput("lista",sID,"hidden","","1","dt","")
   ' sInput   = sinput & ofv.GenerarInput("lista2",sID,"hidden","","1","dt","")
    Formvista = ofv.GenForm("vistar","",sAccion,"","post"," style='display:none;' ",sInput,"si") & "Cambiar a Vista Web"

    Celda  = OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;'  ", "", "", "", "", "", sImagena & "&nbsp;" & sTitulo & "", "si")
    Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ","center onclick='document.all.vistar.submit();' ","20%","","","",Formvista,"si")
    Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;' ","right","20%","","","","&nbsp;" & sTitOrd2 & "Orden:&nbsp;" & sTitOrd & "&nbsp;&nbsp;","si")



    Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



    tablaz = OFV.GenTabla("", "aaa bgcolor='#eeeeee' style='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "100% ", "", "0", "", "1", "2", "", "", Fila, "si")



   sHTML = sHTML & Tablaz

    Fila = ""


   '****** CREACION DE ENCABEZADO DE TITULOS
    Celda = OFV.GenCelda("", "ttx colspan=3", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
    Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=T target=_self ><B><u>&nbsp;TIPO&nbsp;</u></B></A>", "si")


    If smenu = "PORESTADO" Then

        If sID = 4 Then
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Eliminacion", "si")
        Else
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CREACION&nbsp;</u></B></A>", "si")
        End If

    Else

       Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CREACION&nbsp;</u></B></A>", "si")

    End If



    If smenu = "PORFECHA" Then
        If sID3 = "Eliminacion" Then
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Eliminacion", "si")
        ElseIf sID3 = "Vencimiento" Then
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
        ElseIf sID3 = "Publicacion" Then
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
        Else
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
        End If
    ELSEIf smenu = "PORREVISION" Then
        Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
    Else
        Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
        Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
    End If

    If smenu <> "PORESTADO" Then
        Celda = Celda & OFV.GenCelda("", "tTx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=E target=_self ><B><u>&nbsp;ESTADO&nbsp;</u></B></A>", "si")
    Else
        Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Autor", "si")
      '  If sID = 3 Or sID = 4 Then
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Aprobo", "si")
      '  End If
        If sID = 4 Then
            Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Elimino", "si")
        End If
    End If
    Fila = fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

    Response.write sHTML
    sHTML = ""
    Response.flush


'''''''''''''''''''' FIN ENCABEZADOS ''''''''''''''''''''''''''''''''''''''

    nLinea = 0

    fi = CLng(CDate("01-01-1900"))
    PosReg = 0

    xTrColor = "#FFFFFF"




    On Error Resume Next
    If Not rs.EOF Then
     rs.Move (CantRegMover)

     Do While (Not rs.EOF) And (PosReg < CantRegMostrar)




        dEmision  = "..."
        dVigencia = "..."
        sArchivo = "Iconos/" & rs(11)
        If Not IsNull(rs(0)) Then stipo = rs(0)

        If smenu = "PORESTADO" Then
            If sID = 4 Then
                If Not IsNull(rs(5)) Then dEmision = CDate(CLng(rs(7)) + fi)
            Else
                If Not IsNull(rs(5)) Then dEmision = CDate(CLng(rs(5)) + fi)
            End If
        Else
            If Not IsNull(rs(5)) Then dEmision = CDate(CLng(rs(5)) + fi)
        End If

        dEmi = InStr(dEmision, "1900")
        If dEmi > 0 Then dEmision = "..."

        If smenu = "PORFECHA" Then
            If sID3 = "Eliminacion" Then
                If Not IsNull(rs(7)) Then dVigencia = CDate(CLng(rs(7)) + fi)
            ElseIf sID3 = "Vencimiento" Then
                If Not IsNull(rs(16)) Then dVigencia = CDate(CLng(rs(16)) + fi)
            ElseIf sID3 = "Publicacion" Then
                If Not IsNull(rs(24)) Then dVigencia = CDate(CLng(rs(24)) + fi)
            Else
                If Not IsNull(rs(6)) Then dVigencia = CDate(CLng(rs(6)) + fi)
            End If
        ElseIf smenu = "PORREVISION" Then
            If Not IsNull(rs(16)) Then dVigencia = CDate(CLng(rs(16)) + fi)
        Else
            If Not IsNull(rs(6)) Then dVigencia = CDate(CLng(rs(6)) + fi)
        End If

        dpublic = "..."

        If smenu <> "PORFECHA" AND smenu <> "PORREVISION"  Then
           If Not IsNull(rs(22)) Then dpublic = CDate(CLng(rs(22)) + fi)
           dVign = InStr(dpublic, "1900")
           If dVign > 0 Then dpublic = "..."
        end if

        dVign = InStr(dVigencia, "1900")
        If dVign > 0 Then dVigencia = "..."

        If smenu <> "PORESTADO" Then
            If Not IsNull(rs(8)) Then
                If (Not CDate(CLng(rs(16)) + fi) > Date) And (CDate(CLng(rs(16)) + fi) <> CDate("01-01-1900")) Then
                    sEstado = "VENCIDA"
                Else
                    sEstado = rs(8)
                End If
            End If
        Else
            sAutor   = "&nbsp;"
            sAprobo  = "&nbsp;"
            sElimino = "&nbsp;"
            If Not IsNull(rs(12)) Then sAutor   = rs(12)
            If Not IsNull(rs(13)) Then sAprobo  = rs(13)
            If Not IsNull(rs(14)) Then sElimino = rs(14)
        End If

        If Not IsNull(rs(9)) Then nIden = rs(9)
        If Not IsNull(rs(10)) Then sobjeto = rs(10)
        if SESSION("FULLDOCN") > "2" then
           sobjeto = sobjeto & " - " & nIden
        end if
        If smenu <> "PORFECHA" AND smenu <> "PORREVISION"  Then

            If Not IsNull(rs(19)) Then sobjeto = sobjeto & " - " & rs(19)
            If Not IsNull(rs(20)) Then sobjeto = sobjeto & " " & rs(20)

        else

        if smenu="PORFECHA"   THEN
         CALL GETLISTOBJ("E",rS(3),0)
         sobjeto = XGUTABLISDAT
        end if
        END IF



        If Session("USRES") <> "B" Then
            sImagen = OFV.GenImage("", "", "", "", sArchivo, "0", "")
        Else
            sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"
        End If

        sImagen = "<font face='corflaticon' size=3 color='#993333'>4</font>"

        sActcon = "Consultar.asp?Codigo=" & rs("codobjeto") & "&Tipo=" & rs("tipobjeto")
        sActcon = sActcon & "&Perfil=" & rs("perfil") & "&Datos=" & OFV.convertircar(sMenuMas, "&", "*")
        sActcon = sActcon & "&Volver=" & sVolver
        sHref = "&nbsp;"
        sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)

        sLink = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen, "si")

        Celda = OFV.GenCelda("", "ut", "", "3%", "", "", "", sLink, "si")

        smasc = "&nbsp;"
        If smenu <> "PORFECHA"  AND smenu <> "PORREVISION" Then
            If  rs(21) = 1 and  SESSION("UN") <> 1 and  SESSION("EMPMUN") = "S" Then
                smasc = "&nbsp;<font color=#993333  ><font class=btic title='Documento Corporativo' >G</font></font>"
            End If
        End If

        sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, sobjeto, "si")
        sTipoDsc = stipo
        sTipoDsc = "<a class=nt title='solo documentos de tipo " & stipo & "' href=" & starget & "&sorden=" & sorden & "&sfiltro=" & Replace(stipo, " ", "_") & " target=_self ><font color='" & Session("CMFPNIPFT") & "' >" & sTipoDsc & "</font></A>"

        Celda = Celda & OFV.GenCelda("", "ut ", "", "35%", "", "", "", sLink, "si")
        Celda = Celda & OFV.GenCelda("", "ut ", "", "4%", "", "", "", smasc, "si")
        Celda = Celda & OFV.GenCelda("", "ut ", "", "15%", "", "", "","" &  sTipoDsc, "si")
        Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &   fechadorCor(dEmision,""), "si")
        If smenu <> "PORFECHA"  AND smenu <> "PORREVISION" Then
            Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &   fechadorCor(dpublic,""), "si")
        End If



        Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &   fechadorCor(dVigencia,""), "si")

        If smenu <> "PORESTADO" Then
            If sEstado = "VENCIDA" Then
                sInput = "<font color=#880000>" & sEstado & "</font>"
            Else
                sInput = sEstado
            End If
            Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sInput, "si")
        Else
            Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sAutor, "si")
         '   If sID = 3 Or sID = 4 Then
                sAprobo = "<a href=coobjwfldet.asp?Tipobjeto=" & rs("tipobjeto") & "&Codobjeto="
                sAprobo = sAprobo & rs("codobjeto") & " target=_blank><IMG SRC=Imagenes\ICorSInfoUser.jpg title=" & Chr(34)
                sAprobo = sAprobo & " Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & Chr(34) & " border=0></a>"

                sRef     = "coobjwfldet.asp?Tipobjeto=" & rs("tipobjeto") & "&Codobjeto=" & rs("codobjeto")
                sImagenP = "&#0232;"
                sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
                sinput   = "<a class='ft1n' " & sFuncAnc & " href='" & sRef & "' "
                sinput   = sinput  & "' title='Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & sobjeto & "' target='_blank' >" & sImagenP & "</a>"
                sAprobo  =  sinput

                Celda    = Celda & OFV.GenCelda("", "ft", "", "5%", "", "", "", sAprobo, "si")
           ' End If
            If sID = 4 Then
                Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sElimino, "si")
            End If
        End If


        Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F0F0F0"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF

        PosReg = PosReg + 1
        rs.movenext
        nLinea = nLinea + 1
        Response.write sHTML
        sHTML = ""
        If nLinea > 20 Then
            nLinea = 0
            Response.flush
        End If



    Loop

   else

        XFRASE = "<BR><BR>&nbsp;<B><font color=#993333 SIZE=3 FACE='corflaticon' >i</font></B>"
        XFRASE = XFRASE &  "&nbsp;NO SE ENCONTRARON DOCUMENTOS PARA LA BUSQUEDA SOLICITADA<BR><BR>"

        Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", XFRASE, "si")




        Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")


   end if


    sHTML = sHTML & OFV.GenTabla("", "", "", "100%  style={border-style:solid;border-width:1;border-color:#eeeeee;} ", "", "0", "", "-1", "0", "", "", Fila, "si")

    IF SMENU = "PORREVISION" THEN
       sAccion = "COFECREVISION.asp?Menu=" & sMenu & "&Nombre=" & sNombre
    ELSE
       sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombre
    END IF


    '****** BOTONES VOLVER, ANTERIOR Y SIGUIENTE

    Celda   = ofv.BotonVolver(sAccion)
    sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", Listar, Celda)
    'response.write  "<br> Addqry :   "& Addqry
    'response.write "<br> starget:    " & starget
    'response.write "<br> svolver:    " & svolver


    sHTML = sHTML  & "</body></html>"

    If smenu = "PORSECTOR" Then
        Session("sperfil") = Replace(SPerfil, " ", "_")
    Else
       Session("sperfil") = ""
    End If
    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

    Generar = sHTML



  sHTML = Generar

 '' sHTML = oClc.Generar



end if

response.write sHTML
response.flush

function vertexto(nwtexto)

    nwtextoF = nwtexto & ""

    if  nwtextoF <> "" then

        nwtextoA = split(nwtextoF," ")

        nwsep    = ""
        nwtextoF = ""

        for iii2 = 0 to ubound(nwtextoA)

            if len(nwtextoA(iii2)) > 2 then

                nwtextoF =  nwtextoF & nwsep & nwtextoA(iii2)

                nwsep = "%"
            end if
        next


    end if



    vertexto = nwtextoF

end function

%>