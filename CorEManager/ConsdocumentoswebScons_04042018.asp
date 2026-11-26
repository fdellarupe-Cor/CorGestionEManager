<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


  'response.write ("DESDE: "&Request.Form("Desde") & "<br>")
  'response.write ("HASTA: "&Request.Form("Hasta") & "<br>")
  'response.write ("FECHA: "&Request.Form("Fecha") & "<br>")

  'response.write ("DESDEq: "&Request.querystring("desde") & "<br>")
  'response.write ("HASTAq: "&Request.querystring("hasta") & "<br>")
  'response.write ("fechaq: "&Request.querystring("fecha") & "<br>")

  'response.write ("SID: "&Request.querystring("ID") & "<br>")
  'response.write ("SID2: "&Request.querystring("ID2") & "<br>")
  'response.write ("nacta : " & request.querystring("nacta")& "<br>")
  'response.write ("Smenu: "&Request.querystring("menu") & "<br>")
  'response.write ("Session EMPMUN: "& SESSION("EMPMUN") &"")


'response.write request.querystring

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  server.scripttimeout = 600
  response.buffer = true



''''''''''''''
'response.write "REALIZANDO CAMBIOS"

  XXXX = NEWEMPRE("NO")

    volver = OFV.vueltaasp2("../")
    volver2 = OFV.vueltaasp2("")

    sNombre = "Tema - Subtema"

    sID = Request.querystring("ID")

    sitem = OFV.convertircar(Request.querystring("sitem"), "_", " ")
    stitem = OFV.convertircar(Request.querystring("stitem"), "_", " ")
    Stnivel = Request.querystring("stnivel")

   If Request.querystring("stesq") = "" Then
        stesq = "TEMA"
    Else
        stesq = Request.querystring("stesq")
    End If

    volver2 = volver2 & "?ID=" & sID
    volver2 = volver2 & "&sitem=" & OFV.convertircar(sitem, " ", "_")
    volver2 = volver2 & "&stitem=" & OFV.convertircar(stitem, " ", "_")
    volver2 = volver2 & "&stnivel=" & Stnivel & "&stesq=" & stesq
    volver2 = OFV.convertircar(volver2, "&", "*")



     smenu   = Request.querystring("Menu")
     sorden  = Request.querystring("sorden")
     sNombre = request.querystring("Nombre")
sTitOrd=""



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


    ReDim Preserve bConsultar(4)
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "4"
    bConsultar(4) = OFV.AConsultar
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "2"
    bConsultar(2) = OFV.AConsultar
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "1"
    bConsultar(1) = OFV.AConsultar
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "3"
    bConsultar(3) = OFV.AConsultar


	'RESPONSE.WRITE (CHR(13) & smenu & CHR(13))



     If smenu = "PORFECHA" OR smenu = "PORREVISION"  Then
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



        'response.write " ** Sperfil:" & Sperfil & "<br>"
	'response.write " **  SID:" & SID  & "<br>"

If smenu = "PORACTA" Then


         If  Robj = "F"  Then

             sID    = Request.Form("Desde") & ""
             sID2   = Request.Form("Hasta") & ""
             Addqry = "&Desde=" & sID & "&Hasta=" & sID2
         Else
             sID    = Request.querystring("Desde") & ""
             sID2   = Request.querystring("Hasta") & ""
             Addqry = "&Desde=" & sID & "&Hasta=" & sID2
         End If

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
        sMenuMas    = sMenuMas & sID2

        xsid        = Replace(sID, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID  = "") Then  sID  = "1-1-1950"

        xsid        = Replace(sID2, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

   ELSEIf smenu = "PORESTADO" Then


         If  Robj = "F"  Then
             sIDx  = request.form("lista")
             sID    = Request.Form("Desde") & ""
             sID2   = Request.Form("Hasta") & ""
        ''     Addqry = "&Desde=" & sID & "&Hasta=" & sID2

  			 Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&lista=" & sIDx
         Else
         	 sIDx    = request.querystring("lista")
             sID    = Request.querystring("Desde") & ""
             sID2   = Request.querystring("Hasta") & ""
          ''   Addqry = "&Desde=" & sID & "&Hasta=" & sID2
  			 Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&lista=" & sIDx

         End If

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
        sMenuMas    = sMenuMas & sID2

        xsid        = Replace(sID, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID  = "") Then  sID  = "1-1-1950"

        xsid        = Replace(sID2, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

   ELSEIf smenu = "PORREVISION" Then


         If  Robj = "F"  Then

             sID    = Request.Form("Desde") & ""
             sID2   = Request.Form("Hasta") & ""
             Addqry = "&Desde=" & sID & "&Hasta=" & sID2
         Else
             sID    = Request.querystring("Desde") & ""
             sID2   = Request.querystring("Hasta") & ""
             Addqry = "&Desde=" & sID & "&Hasta=" & sID2
         End If

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
        sMenuMas    = sMenuMas & sID2

        xsid        = Replace(sID, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID  = "") Then  sID  = "1-1-1950"

        xsid        = Replace(sID2, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

    ElseIf smenu = "PORFECHA" Then

         If  Robj = "F"  Then

             sID    = Request.Form("Desde") & ""
             sID2   = Request.Form("Hasta") & ""
             sID3   = Request.Form("Fecha") & ""
             Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
         Else
             sID    = Request.querystring("Desde") & ""
             sID2   = Request.querystring("Hasta") & ""
             sID3   = Request.querystring("Fecha") & ""

             Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
         End If

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
        sMenuMas    = sMenuMas & sID2 & "&Fecha=" & sID3




    ElseIf smenu = "PORPRODUCTO" Then

        If  Robj = "F" Then

            sID    = Request.Form("ListaS") & ""
            sID2   = Request.Form("ListaG") & ""
            sID3   = Request.Form("ListaP") & ""
            Addqry = "&ListaS=" & sID & "&ListaG=" & sID2 & "&ListaP=" & sID3
            sID    = Replace(sID, "_", " ")
            sID2   = Replace(sID2, "_", " ")
            sID3   = Replace(sID3, "_", " ")

        Else
            sID    = Request.querystring("ListaS") & ""
            sID2   = Request.querystring("ListaG") & ""
            sID3   = Request.querystring("ListaP") & ""
            Addqry = "&ListaS=" & sID & "&ListaG=" & sID2 & "&ListaP=" & sID3
            sID    = Replace(sID, "_", " ")
            sID2   = Replace(sID2, "_", " ")
            sID3   = Replace(sID3, "_", " ")

        End If

    ElseIf smenu = "PORNROCIRCULAR" Then

        If  Robj = "F" Then

            sID    = Request.Form("ListaG") & ""
            sID2   = Request.Form("ListaN") & ""
            Addqry = "&ListaG=" & sID & "&ListaN=" & sID2
            sID    = Replace(sID, "_", " ")
            sID2   = Replace(sID2, "_", " ")
        Else

            sID    = Request.querystring("ListaG") & ""
            sID2   = Request.querystring("ListaN") & ""
            Addqry = "&ListaG=" & sID & "&ListaN=" & sID2
            sID    = Replace(sID, "_", " ")
            sID2   = Replace(sID2, "_", " ")

        End If

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

            'response.write Addqry

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

        If  Robj = "F" Then

            sID    = Request.Form("buscar") & ""
            Addqry = "&buscar=" & sID

        Else
            sID    = Request.querystring("buscar") & ""
            Addqry = "&buscar=" & sID
        End If

        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&IDENTIFICACION=" & sID



    ElseIf smenu = "PORACTAQ" then




'****

      ElseIf smenu = "PORSECTOR" then
       If  Robj = "F" Then

            SPerfil = Request.Form("perfil") & ""
            sID     = Request.Form("lista") & ""
            sID2    = Request.Form("lista") & ""
            sID3    = Request.Form("buscar") & ""
            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
            SPerfil = Replace(SPerfil, "_", " ")


        Else

            SPerfil = Request.querystring("perfil") & ""
            sID     = Request.querystring("ID") & ""
            sID     = Request.querystring("Lista") & ""
            sID2    = Request.querystring("Lista") & ""
            sID3    = Request.querystring("buscar") & ""
            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
            SPerfil = Replace(SPerfil, "_", " ")


        End If
 ''''
	'response.write "Sperfil:" & Sperfil & "<br>"
	'response.write "SID:" & SID

        sID         = OFV.convertircar(sID, "_", " ")


        If sID = "" Then sID = sID3

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
        sMenuMas    = sMenuMas & sID3 & "&perfil=" & SPerfil


'***


    Else

          ' response.write "aca"
        If  Robj = "F" Then

            SPerfil = Request.Form("perfil") & ""
            sID     = Request.Form("lista") & ""
            sID2    = Request.Form("lista") & ""
            sID3    = Request.Form("buscar") & ""
            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
            SPerfil = Replace(SPerfil, "_", " ")


        Else

            SPerfil = Request.querystring("perfil") & ""
            sID     = Request.querystring("ID") & ""
            if sid="" THEN sID     = Request.querystring("Lista") & ""
            sID2    = Request.querystring("Lista") & ""
            IF SID2="" THEN SID2=Request.querystring("ID") & ""
            sID3    = Request.querystring("buscar") & ""
            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
            SPerfil = Replace(SPerfil, "_", " ")


        End If
 ''''

        sID         = OFV.convertircar(sID, "_", " ")


        If sID = "" Then sID = sID3

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
        sMenuMas    = sMenuMas & sID3 & "&perfil=" & SPerfil


    End If

    Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre

    starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry
    sVolver = "ListadoConsultas.asp?Menu=" & smenu & sMenuAdd

    Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")

    'response.write Addqry

    Set cn = OFV.conectar(OFV.strconn2)

  '''  strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
  '''  strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, "
  '''  strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, "
  '''  strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
  '''  strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,"
  '''  strsql = strsql & "Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,"
  '''  strsql = strsql & "OBJETOS.OBJUN,OBJPRTY.fecpublic "



    strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
    strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, "
    strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, "
    strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
    strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,"
    strsql = strsql & "Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,"
    strsql = strsql & "OBJETOS.OBJUN,OBJPRTY.fecpublic    "
    Strsql = Strsql & ", t.descri, t.descri, t.descri, t.descri, st.descri   "


  '     Strsql = Strsql & " from csrelpath a "
  '     Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto) "
  '     Strsql = Strsql & " inner join objprty c on (a.cobj = c.codobjeto) "
  '     Strsql = Strsql & " inner join estados e on (c.estado = e.objid) "
  '     Strsql = Strsql & " inner join tipobjeto f on (b.tipobjeto = f.tipobjeto) "
  '     Strsql = Strsql & " inner join tema      t on (a.tema = t.codtema) "
  '     Strsql = Strsql & " inner join prodserv  st on (a.stema = st.codprodserv) "
  '     Strsql = Strsql & " left  join objinfo g on (b.codobjeto = g.codobjeto) "


    addfrom = addfrom & ", csrelpath a, tema t, prodserv  st"
  ''''  addfrom = addfrom & ", objinfo g "

    wherefrom = " and (a.cobj = OBJETOS.codobjeto) "
    wherefrom = wherefrom & " and (a.tema = t.codtema) "
    wherefrom = wherefrom & " and (a.stema = st.codprodserv) "
 ''''   wherefrom = wherefrom & " and (OBJETOS.codobjeto = g.codobjeto) "


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




        Case "PORESTADO"

            ssEstado = ""

            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
            strsql = strsql & " where  "
            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
            strsql = strsql & "and OBJPRTY.estado = " & sID
            sTitulo = "Estado: " & sID
            sqle = "Select * from estados where objid =  " & sidx
      ''  response.write "<br><br><br>" & sqle & "<br>"
            Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
            If Not rsx.EOF Then
            	 sTitulo = "Estado: " & rsx(1)
               ssEstado = rsx(1)
            End if



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

            'Response.write strsql
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
                   strsql = strsql & " AND OBJETOS.TSEGMENTO = " & CDbl("0" &sID3) & " "
                End If

                If UCase(sID & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
                   strsql = strsql & " AND ((OBJETOS.codobjeto = OBJTSEGAGR.codobjeto)  AND ( "
                End If

            If UCase(sID & "") <> UCase("nsnc") Then
                strsql = strsql & " (OBJTSEGAGR.TSEGMENTO = " & CDbl("0" & sID) & " ) "
            End If
            If UCase(sID2 & "") <> UCase("nsnc") Then
                   If UCase(sID & "") <> UCase("nsnc") Then strsql = strsql & " AND "
                   strsql = strsql & " (OBJTSEGAGR.TAGRUPA = " & CDbl("0" & sID2) & " ) "
                End If

                If UCase(sID & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
                   strsql = strsql & " ) )"
                End If


            xinstrum = ""
            If UCase(sID & "") <> UCase("nsnc") Then
                Sqlx = "SELECT descri from tiposegmento where objid=" & CDbl("0" & sID)
                Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
                If Not rsx.EOF Then xinstrum = rsx(0) & ""
                Call OFV.cerrarconsulta(rsx)
            End If


            xinstrum2 = ""
            If UCase(sID2 & "") <> UCase("nsnc") Then
                Sqlx = "SELECT descri from tipoAGRUPA where objid=" & CDbl("0"&  sID2)
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
                   strsql = strsql & " AND OBJETOS.codiden = '" & CStr( sID2 & "") & "' "
                End If

                If UCase(sID & "") <> UCase("nsnc") Then
                    strsql = strsql & " AND ((OBJETOS.codobjeto = OBJTSEGAGR.codobjeto)  AND (OBJTSEGAGR.TAGRUPA = " & CDbl( "0" & sID) & " ) ) "
                End If
                 'Response.write sID & "<br>"

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
                   strsql = strsql & " on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO " & addfrom
          '     END IF


            Strsql = strsql & " WHERE  "

            strsql = strsql & " ( "

			'Response.write sID & "<br>"
			'Response.write sID2 & "<br>"
            If sID2 <> "" Then
                strsql = strsql & "  ((OBJPRTY.instrumen=" & CDbl("0" & sID) & ")"
            Else
                strsql = strsql & "  (OBJPRTY.instrumen=" & CDbl("0" & sID) & " or OBJPRTYADV.instrua=" & CDbl("0" & sID) & ")  "
            End If

            If sID2 <> "" Then
                strsql = strsql & " AND (OBJPRTY.instid ='" & sID2 & "')) OR "
                strsql = strsql & " (OBJPRTYADV.instrua = 0" & sID & " and OBJPRTYADV.instid ='" & sID2 & "')"
            End If

            strsql = strsql & " ) "

            xinstrum = ""
            Sqlx = "SELECT descri from instrumen where objid=" & CDbl("0" & sID)
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

        Case "PORSECTOR"



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

                    strsql = strsql & " and(( objproced.codobjeto=OBJETOS.codobjeto "
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

''     if  SPerfil <> "nsnc" then sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & SPerfil
       if  SPerfil <> "nsnc" then

'' AGREGADO 26052014
           sTitulo = sTitulo & "&nbsp;"

           Strsqle = "Select descri from tipoperfil where Tipoperfil ='" & SPerfil & "' "
		       Set rsxx = OFV.crearconsultaEx(Strsqle, cn, 1, "")
		       if not rsxx.eof then
			        sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & rsxx(0)
		       else
			        sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & "TODOS"
		       end if

           call ofv.cerrarconsulta(rsxx)

       else
       	   sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & "TODOS"
       end if

'' AGREGADO 26052014 Fin

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
            strsql = "select * from csporfechas " & addfrom

           ' cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
            'chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

            cdesde = CDate(  sID)
            chasta =CDate( sID2)

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

        Case "PORREVISION"
            strsql = "select * from csporfechas " & addfrom
            cdesde = CDate(  (sID)  + CLng(CDate("01-01-1900")))
            chasta = CDate( (sID2)  +   CLng(CDate("01-01-1900")  ))

            sid = CLng(CDate(sID)) 'CDate(  cdate(sID)  + CLng(CDate("01-01-1900")) )
            sid2 =CLng(CDate(sID2))' CDate( cdate(sID2) + CLng(CDate("01-01-1900")) )
           ' response.write sid
            strsql = strsql & " WHERE  xfecvto between '" & sid & "' and '" & sid2 & "' "

            sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & cdesde & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & chasta




	Case "PORACTA"

			cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
            chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))


			If  Robj = "F"  Then

				 sID    = Request.Form("Desde") & ""
				 sID2   = Request.Form("Hasta") & ""
				 sID3   = Request.Form("Fecha") & ""
				 nacta  = Request.Form("nacta")
				 Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3

			 Else
				 sID    = Request.querystring("Desde") & ""
				 sID2   = Request.querystring("Hasta") & ""
				 sID3   = Request.querystring("Fecha") & ""
				 nacta  = Request.querystring("nacta") & ""

				 Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3


			 End If


			sMenuAdd    = ""
			sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
			sMenuMas    =  sMenuMas & sID2 & "&Fecha=" & sID3


			if  nacta <> "" then
				buscaacta = " Objprty.nroacta = '" & nacta & "'"
			else
				buscaacta = "Objprty.nroacta is not null"
			end if


           		 cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
          		 chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))


			strsql = "select DISTINCT csporfechas.*, Objprty.fecacta,objetos.descri,tipobjeto.descri  from csporfechas join Objprty on csporfechas.codobjeto = Objprty.codobjeto, Estados  " & addfrom
			strsql = strsql & " where " & buscaacta & " and Objprty.fecacta between '" & cdesde & "' and '" & chasta & "' "

			sTitulo = "Por acta:  " & nacta &"&nbsp;&nbsp; Desde: "& sID & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & sID2

    End Select



    strsql = strsql & sverestado & wherefrom

    sTitOrd  = ""
    sTitOrd2 = ""

    If smenu = "PORFECHA" or smenu = "PORREVISION" Then

        If Sfiltro <> "TODOS" Then
           sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
           strsql  = strsql & " and csporfechas.descri='" & Sfiltro & "' "
        End If

        Select Case sorden

        	      Case "N"
                        sTitOrd = sTitOrd & "Nombre"
                        strsql  = strsql & " order by tipobjeto.descri, csporfechas.descri"
                Case "T"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by  tipobjeto.descri, csporfechas.descri"
                Case "C"
                        sTitOrd = sTitOrd & "Creacion"
                        strsql  = strsql & " order by  tipobjeto.descri, csporfechas.descri"
                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by  tipobjeto.descri, csporfechas.descri"
                Case "P"
                        sTitOrd = sTitOrd & "Publica"
                        strsql  = strsql & " order by  fecpublic, csporfechas.descri"
                Case "V"
                        sTitOrd = sTitOrd & "Vigencia"
                        strsql  = strsql & " order by  fecvigen, csporfechas.descri"
                Case "A"
                        sTitOrd = sTitOrd & "Autor"
                        strsql  = strsql & " order by  autor, csporfechas.descri"

        End Select



	ElseIf smenu = "PORACTA" THEN
		strsql  = strsql & " order by csporfechas.NROACTA, Objprty.FECACTA"



    Else
        If Sfiltro <> "TODOS" Then
           sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
           strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
        End If

        Select Case sorden
                Case "N"
                        sTitOrd = sTitOrd & "Nombre"
                        strsql  = strsql & " order by OBJETOS.descri,tipobjeto.descri "
                Case "T"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by tipobjeto.descri, objetos.descri"
                Case "C"
                        sTitOrd = sTitOrd & "Creacion"
                        strsql  = strsql & " order by objprty.fecemis,tipobjeto.descri"
                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by estdescri,objetos.descri,tipobjeto.descri "
                Case "P"
                        sTitOrd = sTitOrd & "Publica"
                        strsql  = strsql & " order by  fecpublic, csporfechas.descri"
                Case "V"
                        sTitOrd = sTitOrd & "Vigencia"
                        strsql  = strsql & " order by  fecvigen, csporfechas.descri"
                Case "A"
                        sTitOrd = sTitOrd & "Autor"
                        strsql  = strsql & " order by  autor, csporfechas.descri"
        End Select
    End If

  'response.write "<br><br><br>" & sTitOrd & "<br>" & chr(13)
  'response.write (smenu) & "----------" & sorden
 'response.write "<br><br><br>" & strsql & "<br>"

    'Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

      sHTML      = ofv.FormHeader(session("FormN"))
      xs = "<script language=JavaScript>" & Chr(13)
      xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	    xs = xs & " var h=570;w=790 ;"
		if  session("APLMDCLI")="BC" THEN
	      xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
		END IF

	    xs = xs & " myWin= open(Datos"
	    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & ","' & Chr(34)
	    'xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
	    xs = xs &"  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'" '& CHR(34)
	    xs = xs & " ); " & Chr(13) & " } " & Chr(13)


    xs = xs & "function Axn(c1,c2,c3) {" & Chr(13)

    xs = xs & " var mydat= 'Consultar.asp?DOCMODAL=CL&Codigo=' + c1 + '&Tipo=' + c2 + '&Perfil=' + c3 + '&Datos=&Volver=" & volver2 & "'; "
    xs = xs & " var h=570;w=790 ;"
   		if  session("APLMDCLI")="BC" THEN
   	      xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
		END IF
    xs = xs & " myWin= open(mydat"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," '& Chr(34)
   ' xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
    xs = xs &"  'menubar=no,toolbar=no,width='+ w +',height='+ h +',top=0,left=0,resizable=yes,status=yes'" '& CHR(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)


    xs = xs & "</script>" & Chr(13)

    sHTML = sHTML & xs


  Fila    = ""
  Celda   = ""
  sHref   = ""

  dim tabprocesos (500,500,20)
  dim tabprocord  (500,5)


  ttabprocesos = 0


  Fila = ""


   sHTML = sHTML & "<center>"


    Response.write sHTML
    sHTML = ""
    Response.flush

    Set cn = OFV.conectar(OFV.strconn2)

    xTrColor = "#FFFFFF"



'''''''''''''' indice normativo     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  TABLAB = ""
  TABLAB = "<DIV id='DTindex' STYLE='float: left; z-index:1;overflow:NONE;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"


    'sInputT = "&nbsp;&nbsp;&nbsp;<B><font color='#0000AA' size=2 >&#0124;</font></B>&nbsp;&nbsp;<font color='#000000' size=1 ><B>" & sTitulo & "</B></font>&nbsp;"
	'sInputT = "a ver que pone!!!"
    celda  = ""
    'Celda      = Celda & ofv.GenCelda("","vti STYLE='cursor:pointer;color:#990000;font-size:9pt;' ","","","","","","Busqueda&nbsp;por&nbsp;" & sInputT,"si")
''	Celda      = Celda & ofv.GenCelda("","vti STYLE='cursor:pointer;color:#990000;font-size:9pt;' ","","","","","", "","si")
  '''  Celda      = Celda & ofv.GenCelda("","UT  ","","70%","","","",sInputT,"si")

 ''   FILA     = ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","","","",Celda,"si")

 '' TABLAB = TablaB & ofv.GenTabla("","aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ","","98%","","0","","-2","0","","",Fila,"si")


    sNombre=replace(sNombre," ","_")
    Select Case smenu

    	Case "PORACTA"

	 	sAccion = "ListadoConsultas.asp?Menu=" & sMenu & "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombre & "&Desde="& Sid &"&Hasta="& Sid2 &"&Fecha="& Sid3
    		sAccion = sAccion &"&nacta="& nacta


    	Case "PORESTADO"
    	adc=request.querystring
    	sAccion = "ListadoConsultas.asp?xx=" & Addqry ' & sMenu & "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=10&SORDEN=N" & "&Nombre=" & sNombre & "&Desde="& Sid &"&Hasta="& Sid2 &"&Fecha="& Sid3

    	Case "PORFECHA"

	        sAccion = "ListadoConsultas.asp?xx=" & Addqry ' "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombre & "&Desde="& Sid &"&Hasta="& Sid2 &"&Fecha="& Sid3

    	Case "PORREVISION"

    		sAccion = "ListadoConsultas.asp?xx=" & Addqry ' "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombre & "&Desde="& Sid &"&Hasta="& Sid2 &"&Fecha="& Sid3

      	Case "PORSECTOR"

         	sAccion = "ListadoConsultas.asp?xx=" & Addqry ' "&lista=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombre & "&Perfil="& Sperfil

         Case "PORCOMPONENTE"
             'response.write request.form & "<br>wwwwwwwwwww" & request.querystring
            sid=replace(sid," ","_")
         	sAccion = "ListadoConsultas.asp?xx=" & Addqry ' "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombre & "&Perfil="& Sperfil

         Case "PORINSTRUMENTO"
            sid=replace(sid," ","_")
	 	   sAccion = "ListadoConsultas.asp?xx=" & Addqry ' "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombre & "&Desde="& Sid &"&Hasta="& Sid2 &"&Fecha="& Sid3

          Case "PORTEXTO"

	        sAccion = "ListadoConsultas.asp?xx=" & addqry


	 Case "PORIDENTIFICACION"

		sAccion = "ListadoConsultas.asp?xx=" & addqry


	Case "PORNROCIRCULAR"


         	sAccion = "ListadoConsultas.asp?xx=" & addqry


	Case "PORPRODUCTO"


	      	'sAccion = "ListadoConsultas.asp?xx=" & addqry
	      	sAccion = "ListadoConsultas.asp?xx=" & Addqry ' "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombre & "&Desde="& Sid &"&Hasta="& Sid2 &"&Fecha="& Sid3




    End Select

'response.write sAccion
    'sAccion = "ListadoConsultas.asp?Menu=" & sMenu & "&id=" & sID  & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
     saccion=saccion & "&Scal=S"
   ' sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre
    sInput   = ofv.GenerarInput("","Vista lista","submit","","10","bt","")
    if  smenu <> "PORACTA" AND smenu <> "PORPRODUCTO" anD smenu <> "PORFECHA"  AND smenu <> "PORREVISION" and smenu<> "PORSECTOR" and smenu<> "PORINSTRUMENTO" and smenu<> "PORIDENTIFICACION" then
    sInput   = sinput & ofv.GenerarInput("lista",sID,"","","1","dt","")
    sInput   = sinput & ofv.GenerarInput("lista2",sID,"","","1","dt","")
    end if
    Formvista = ofv.GenForm("vistar","",sAccion,"","post"," style='display:none;' ",sInput,"si") & "Cambiar a Lista"
    Celda  = OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;'  ", "", "", "", "", "", sTitulo , "si")

    If smenu = "PORESTADO" OR  smenu = "PORREVISION"  OR smenu = "PORSECTOR"  OR smenu ="PORCOMPONENTE" OR smenu="PORACTA" OR  smenu="PORFECHA" or smenu = "PORTEXTO" OR smenu = "PORINSTRUMENTO" or smenu="PORIDENTIFICACION" or smenu="PORNROCIRCULAR" or smenu="PORPRODUCTO" Then
       Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ","center onclick=""document.getElementById('vistar').submit();"" ","20%","","","",Formvista,"si")
       'Celda  = Celda & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ","center onclick='document.getElementById('vistar').submit();' ","20%","","","",Formvista,"si")
    end if
    Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;' ","right","20%","","","", sTitOrd2 & "Orden:&nbsp;" & sTitOrd & "&nbsp;&nbsp;","si")
    Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    tablaz = OFV.GenTabla("", "aaa bgcolor='#eeeeee' style='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "100% ", "", "0", "", "1", "0", "", "", Fila, "si") & "<br>"

  TABLAB = Tablaz & TablaB

  TABLAB = TablaB & "</DIV>"

           sHTML = sHTML  &   TABLAB

   Fila = ""

'' AGREGADO 26052014

    if smenu ="PORESTADO" then
    	 if ssEstado = "VIGENTE" then                     'pregunto si es estado vigente para obtener los totales por tipo'
          strsql = " SELECT max(TIPOBJETO.descri), sum(1) as total "
          strsql = strsql & " FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados , csrelpath a, tema t, prodserv  st "
          strsql = strsql & " where  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and OBJPRTY.estado = Estados.objid and OBJETOS.codobjeto = OBJPRTY.codobjeto "
          strsql = strsql & " and OBJPRTY.estado = 3 and Objetos.objnivel >= 3 And Estados.objid In (9,1,6,2,5,3,4)  and Objetos.OBJUN = 1  and TIPOBJETO.TIPOCPT = 'D' "
          strsql = strsql & " and (OBJETOS.APLIC = 'EM' or OBJETOS.APLIC = 'GLB')  and (a.cobj = OBJETOS.codobjeto)  and (a.tema = t.codtema) "
          strsql = strsql & " and (a.stema = st.codprodserv) "
          strsql = strsql & " group by  TIPOBJETO.descri "
         ' response.write " strsql " & strsql  & "<br>"
          Set rsx1 = OFV.crearconsultaEx(strsql, cn, 1, "")

          if NOT rsx1.eof then

	'' Titulo de Totales por Tipo -unica vez-

	           Celda = ""
	           Celda  = OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "20%", "", "", "", "&nbsp;" & "TOTALES POR TIPO" & "", "si")
	           Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
             tablac = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "60% ", "", "0", "", "1", "2", "", "", Fila, "si")
             sHTML = sHTML & Tablac
             Fila = ""

  '' Rutina para mostrar los totales
  '' response.write " count "    & rsx1.fields.count   &"<br>"

             iii = 0
             iii3 = 0
             for iii = 1 to 1      '9 step 3

        '     response.write " &rsx1 " & rsx1(0) & "<br>"
        '     response.write " &iii "  & iii     & "<br>"
        '     response.write " Xrsx1X " & Hxcod & " " & Hxtip & " " & iii &"<br>"

               Fila  = ""
               do while not rsx1.eof
                  Celda = ""
                  for iii3=1 to 3
                  IF  rsx1.eof then
				      exit for 'RSX1.MOVENEXT

                  end if
                      Hxcod = rsx1(0)   'descripcion
                      Hxtip = rsx1(1) & ""  'cantidad

           '' response.write  rsx1(0) &"<br>"

                      Celda  = Celda & OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "20%", "", "", "", "&nbsp;" & Hxcod & ": " & Hxtip & "", "si")

                         RSX1.MOVENEXT

                   next

                   Fila   = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               loop

             next

          end if
       end if

          tablac = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "60% ", "", "0", "", "1", "2", "", "", Fila, "si")
          sHTML = sHTML & Tablac
          Fila = ""

    end if

'' AGREGADO 26052014 FIN

      xsclase =" class=istxz " ' " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "

      XTABLAN = ""


      iproc   = 0


      iprdoc   = 0

      dFechai = clng(date()) - clng(cdate("01-01-1900"))




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''   tabprocesos (iproc,iprdoc,1) = sitem                        'nombre    2
            ''   tabprocesos (iproc,iprdoc,2) = rs(0)                        'codigo    3
            ''   tabprocesos (iproc,iprdoc,3) = sPerfil                      'perfil    15
            ''   tabprocesos (iproc,iprdoc,4) = "Documento normativo"        'copete    23
            ''   tabprocesos (iproc,iprdoc,5) = rs(6)                        'fec.vigencia  6
            ''   tabprocesos (iproc,iprdoc,6) = 0                            'corporativo 0-no 1-si
            ''   tabprocesos (iproc,iprdoc,7) = stitem                       'stitem
            ''   tabprocesos (iproc,iprdoc,9) = rs(19)                        'objimagen    24
            ''   tabprocesos (iproc,iprdoc,10) = rs(20)                        'objurldsc   25
            ''   tabprocesos (iproc,iprdoc,11) = rs(21)                        'tema        26
            ''   tabprocesos (iproc,iprdoc,12) = rs(22)                        'subtema     27



  '''  strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
  '''  strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, "
  '''  strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, "
  '''  strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
  '''  strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,"
  '''  strsql = strsql & "Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,"
  '''  strsql = strsql & "OBJETOS.OBJUN,OBJPRTY.fecpublic, g.objcopete,    "
  '''  Strsql = Strsql & "g.objimagen, g.objurldsc, t.descri, st.descri   "


  '     Strsql = "select a.cobj, B.DESCRI, b.tipobjeto, e.estado,"
  '     Strsql = Strsql & " b.tipobjeto, c.fecemis, c.fecvigen, c.fecbaja, "
  '     StrSql = StrSql & " c.autor, c.aprobo, c.elimino, c.fecvto, b.OBJUN, b.indobj, "
  '     Strsql = Strsql & "  f.descri as tobjdescri , f.perfil, f.TIPOCPT, g.objcopete,    "
  '     Strsql = Strsql & "  c.fecpublic, g.objimagen, g.objurldsc, t.descri, st.descri   "
  '     Strsql = Strsql & " from csrelpath a "
  '     Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto) "
  '     Strsql = Strsql & " inner join objprty c on (a.cobj = c.codobjeto) "
  '     Strsql = Strsql & " inner join estados e on (c.estado = e.objid) "
  '     Strsql = Strsql & " inner join tipobjeto f on (b.tipobjeto = f.tipobjeto) "
  '     Strsql = Strsql & " inner join tema      t on (a.tema = t.codtema) "
  '     Strsql = Strsql & " inner join prodserv  st on (a.stema = st.codprodserv) "
  '     Strsql = Strsql & " left  join objinfo g on (b.codobjeto = g.codobjeto) "

  '     StrSql = StrSql & " where "
  '     StrSql = StrSql & " (b.objun =  " & session("UN") & " or "
  '     StrSql = StrSql & "  b.objun =  1 )  "
  '     StrSql = StrSql & " and a.stema = " & sID & " "
  '     StrSql = StrSql & " and c.estado = 3 "
  '     StrSql = StrSql & " and f.TIPOCPT = 'D' "
  '     StrSql = StrSql & " and f.filtro = 'S' "
  '     StrSql = StrSql & " and (b.aplic = 'GLB' or b.aplic = '" & wkaplic & "') "

  '     StrSql = StrSql & " order by b.tipobjeto, b.indobj,  b.descri, e.estado "



    sNuevoTipo = "ESQUEMAS"
    nNivel = Stnivel + 1
    Session("sID2") = sID

   'response.write strsql & "<br>" & chr(13)

    sVino = "N"

    tipoant = "NO"
    canttipo = 0
   strsql=session("sqlcons")
  'response.write strsql


  'inicio modif Mariano 26/09/2013
   if request.querystring("scal")="S" then
        strsql=session("sqlcons")
      else
         session("sqlcons")=strsql
    end if


  'Fin Inicio Modif Mariano 26/09/2013





'RESPONSE.WRITE smenu &"aa"


    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If rs.EOF Then

       sVino = "N"


    Else



        sVino = "N"

        nlinea = 0
'RESPONSE.WRITE STRSQL
        'Do Until rs.EOF
        rs.Move (CantRegMover)
		     Do While (Not rs.EOF) And (nlinea < CantRegMostrar)

             sVinook = "N"
            ''' if  cstr(rs(12) & "") = cstr(session("UN") & "") then
   '' response.write "<br><br>entro"
                 sVino   = "S"
                 sVinook = "S"
           '''  end if


          if  sVinook = "S" then

                stitem  = rs(4)
                sTipoE  = rs(0)
                 sPerfil = rs(15)
                sNIVE   = "D"
                IF session("APLMDCLI")="BC" THEN
                sitem =	sTipoE & ": <b>" & rs(1) &"&nbsp;-&nbsp;</b> " & Capitalize(rs(2) & "")
                else
                sitem = sTipoE & ": " & Capitalize(rs(2) & "")
                end if

               if SESSION("FULLDOCN") > "2" then
                      NROVER="("
                    If smenu <> "PORFECHA" AND smenu <> "PORREVISION"  Then
                       If Not IsNull(rs(19)) Then NROVER =NROVER & rs(19)
                       If Not IsNull(rs(20)) Then NROVER =  NROVER & "." &  UCASE(rs(20)&"")
				        sitem = sitem & " - "  & " " & UCASE(NROVER &"") & ") (" & UCASE(RS(17)&"") &")"
				      else
				       if smenu="PORFECHA"   THEN
					            CALL GETLISTOBJ("E",rS(3),0)
					            NROVER =  XGUTABLISDAT
              	       end if

                      '' sitem = sTipoE & ": " &  Capitalize(rs(2) & "")
				       sitem =  sitem & nrover  & " (" & UCASE(RS("estdescri")&"") &")"

				      end if
				  '     NROVER=NROVER & "</B>"
        		end if



             '   stitem  = rs(4)
             '   sTipoE  = rs(14)
             '   sPerfil = rs(15)
             '   sNIVE   = rs(16)

             '   sitem = Capitalize(rs(1))


            if tipoant <> sTipoE then

               wspleftb = ""

               if tipoant <> "NO" then


               else

                  wspleftb = "border-left-style:solid;border-left-width:1;border-left-color:#aa0000;"

               end if

               tipoant = sTipoE

               iproc = iproc + 1

               tabprocesos (iproc,0,0) = sTipoE
               tabprocesos (iproc,0,1) = 0


               tabprocord  (iproc,0)   = iproc
               tabprocord  (iproc,1)   = 0

               ttabprocesos = iproc

               iprdoc   = 0


            end if

                 wspwkcopete = ""
                 wspwkimagen = ""
                 wspwkurldsc = ""

                 strsql = "SELECT g.objcopete, g.objimagen, g.objurldsc   "
                 Strsql = Strsql & " from objinfo g where g.codobjeto = 0" & rs(3)

   '              response.write strsql & "<br>"

                 Set rsxx1 = OFV.crearconsultaEx(strsql, cn, 1, "")
                 if not  rsxx1.eof then
                    wspwkcopete = rsxx1(0) & ""
                    wspwkimagen = rsxx1(1) & ""
                    wspwkurldsc = rsxx1(2) & ""
                 end if



            ''   tabprocesos (iproc,iprdoc,1) = sitem                        'nombre    2
            ''   tabprocesos (iproc,iprdoc,2) = rs(0)                        'codigo    3
            ''   tabprocesos (iproc,iprdoc,3) = sPerfil                      'perfil    15
            ''   tabprocesos (iproc,iprdoc,4) = "Documento normativo"        'copete    23
            ''   tabprocesos (iproc,iprdoc,5) = rs(6)                        'fec.vigencia  6
            ''   tabprocesos (iproc,iprdoc,6) = 0                            'corporativo 0-no 1-si
            ''   tabprocesos (iproc,iprdoc,7) = stitem                       'stitem
            ''   tabprocesos (iproc,iprdoc,9) = rs(19)                        'objimagen    24
            ''   tabprocesos (iproc,iprdoc,10) = rs(20)                        'objurldsc   25
            ''   tabprocesos (iproc,iprdoc,11) = rs(21)                        'tema        26
            ''   tabprocesos (iproc,iprdoc,12) = rs(22)                        'subtema     27


               iprdoc = iprdoc + 1

               tabprocesos (iproc,0,1) = iprdoc

               tabprocord  (iproc,1)   = iprdoc


               tabprocesos (iproc,iprdoc,1) =  sitem                        'nombre

               tabprocesos (iproc,iprdoc,2)  = rs(3)                        'codigo
               tabprocesos (iproc,iprdoc,14) = rs(4)                        'tipo

               tabprocesos (iproc,iprdoc,3) = sPerfil                      'perfil


               if wspwkcopete = "" then
                  tabprocesos (iproc,iprdoc,4) = "Documento normativo"        'copete
               else
                  tabprocesos (iproc,iprdoc,4) = Capitalize(ltrim(wspwkcopete))      'copete
               end if


				IF smenu = "PORACTA" THEN
					tabprocesos (iproc,iprdoc,5) = rs("fecacta")                        'fec.ACTA
				ELSE
					tabprocesos (iproc,iprdoc,5) = rs("fecpublic")                        'fec.vigencia
					if clng("0" & rs("fecpublic")) = 0 then tabprocesos (iproc,iprdoc,5) = rs("fecemis")       'fec.emis
				END IF


             '  If clng("0" & rs(12)) = 1 and  SESSION("UN") <> 1 and  SESSION("EMPMUN") = "S" Then
             '     tabprocesos (iproc,iprdoc,6) = 1                            'corporativo 0-no 1-si
             '  Else
                  tabprocesos (iproc,iprdoc,6) = 0                            'corporativo 0-no 1-si
             '  End If

               tabprocesos (iproc,iprdoc,7)  = stitem                        'stitem
               tabprocesos (iproc,iprdoc,8)  = rs("fecpublic")                        'fec.public
               if clng("0" & rs("fecpublic")) = 0 then tabprocesos (iproc,iprdoc,8) = rs("fecemis")       'fec.emis
               tabprocesos (iproc,iprdoc,9)  = wspwkimagen                   'objimagen
               tabprocesos (iproc,iprdoc,10) = wspwkurldsc                   'objurldsc

			   IF smenu = "PORACTA" THEN
					tabprocesos (iproc,iprdoc,11) = "NRO ACTA"                    'descripcion
					tabprocesos (iproc,iprdoc,12) = rs("nroacta")		'Numero de acta
			   ELSEIF (smenu = "PORFECHA") THEN
			        ' response.write "22222222222222222222222222"
					tabprocesos (iproc,iprdoc,11) = ""'rs(26)                        'tema
					tabprocesos (iproc,iprdoc,12) = ""'rs(27)                        'subtema
			   ELSEIF (smenu = "PORREVISION") THEN
					tabprocesos (iproc,iprdoc,11) =""' rs(92)                        'tema
					tabprocesos (iproc,iprdoc,12) = ""'rs(98)                        'subtema
			   ELSE
					tabprocesos (iproc,iprdoc,11) =""' rs(26)                        'tema
					tabprocesos (iproc,iprdoc,12) = ""'rs(27)                        'subtema
			   END IF
			   tabprocesos (iproc,iprdoc,13) = rs("fecvigen")                         'fec.vigencia
			   if clng("0" & rs("fecvigen")) = 0 then tabprocesos (iproc,iprdoc,13) = rs("fecemis")       'fec.emis


            end if

            rs.movenext
            nlinea = nlinea + 1

            if nlinea=499 then exit do
        Loop





    End If

    if  sVino = "N" then

        Celda   = ""

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        'sLink   = "No se encuentran documentos para esta busqueda."
		sLink = "<BR><B><font color=#993333 SIZE=1 FACE='corflaticon' >i</font></B>"
        sLink = sLink &  "&nbsp;NO SE ENCONTRARON DOCUMENTOS PARA LA BUSQUEDA SOLICITADA"
        Celda = Celda & ofv.GenCelda("","ut colspan=2 ","","","","","",sLink,"si")


        Fila    = Fila & ofv.GenRow("", "aaa bgcolor=" & xTrColor & " ","", "", "22", "", "", Celda, "si")

        sHTML = sHTML & "<table>" & Fila & "</table>"

    else

     ''''   response.write "cant proc: " & ttabprocesos & "<br>"


       if ttabprocesos > 1 then

          for ii=1 to ttabprocesos

              for ii2=1 to (ttabprocesos - ii)

                  if ii2 < ttabprocesos then
                     if (tabprocord  (ii2,1) < tabprocord  (ii2+1,1)) or ( (tabprocord  (ii2,1) = tabprocord  (ii2+1,1)) and (tabprocord  (ii2,0) > tabprocord  (ii2+1,0)) ) then

                        wspwkind =  tabprocord  (ii2,0)
                        wspwkpro =  tabprocord  (ii2,1)

                        tabprocord  (ii2,0) = tabprocord  (ii2+1,0)
                        tabprocord  (ii2,1) = tabprocord  (ii2+1,1)

                        tabprocord  (ii2+1,0) = wspwkind
                        tabprocord  (ii2+1,1) = wspwkpro



                     end if

                  end if


              next


          next

       end if

      xsclase = " CLASS=istxz " '" CLASS=VTI STYLE='background-color:#aaaaff;color:#000000;font-size:8pt;' "


        for ii=1 to ttabprocesos

			'if smenu <> "PORACTA" then
			'	iproc = tabprocord  (ii,0)
			'else
				iproc = ii
		'	end if


          ''     tabprocesos (iproc,0,0)  '' descripcion del tipo de documento


          ''     tabprocesos (iproc,0,1)  '' cantidad de documentos


           for iprdoc=1 to tabprocesos (iproc,0,1)


            XTABLAN = XTablaN & chr(13)
            XTABLAN = XTablaN & "<div id='DTNOVINT" & ii & "-" & iprdoc & "' style='float: left; display: inline;  "
            XTABLAN = XTablaN & "border-style:solid;border-width:4;border-color:#ffffff;' >"


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


                 xtmsgIMG = tabprocesos (iproc,iprdoc,9) & ""
                 xtmsgVID = tabprocesos (iproc,iprdoc,10) & ""
                 xtmcod   = tabprocesos (iproc,iprdoc,2) & ""

                 xtmdtema   = tabprocesos (iproc,iprdoc,11) & ""
                 xtmdstema  = tabprocesos (iproc,iprdoc,12) & ""



                If tabprocesos (iproc,iprdoc,6) = 1 Then
                   smasc = "&nbsp;&nbsp;&nbsp;<font color=#993333 ><font class=btic title='Corporativo' >G</font></font>"
                Else
                   smasc = "&nbsp;"
                End If


                Celda   = ""

                sImage  = "<font face='corflaticon' size=4 color='#005500' >&#0157;</font>"
                sHref    = "&nbsp;"
                sFuncion = " onclick=" & Chr(34) & " Axn('" & tabprocesos (iproc,iprdoc,2) & "','" & tabprocesos (iproc,iprdoc,7) & "','" & tabprocesos (iproc,iprdoc,3) & "'); return false;" & Chr(34)

                sLink1    = "&nbsp;<font color='#AA0000' size=2 >&#0149;</font>&nbsp;"
                sLink    = "&nbsp;&nbsp;" & tabprocesos (iproc,iprdoc,1)


                       ' if session("APLMDCLI")="BC" THEN
                      '   xtmsgtit =  "<b>" & tabprocesos (iproc,iprdoc,2) & "&nbsp;-&nbsp;</b>" & Capitalize(tabprocesos (iproc,iprdoc,1))
                     '   ELSE
                     'response.write iproc & "   iproc"
                     'response.write iprdoc & "   iprdoc</br>"
                     'response.write sitem

                         xtmsgtit =   Capitalize(tabprocesos (iproc,iprdoc,1))
                     '   END IF


                         if (dFechai - tabprocesos (iproc,iprdoc,8)) < 31 then


                            xtmsgtit = "<B><font color='#ffffff' STYLE='background-color:#ff0000;' size=1 >&nbsp;nuevo&nbsp;</font></B>&nbsp;&nbsp;" & xtmsgtit

                         end if

'''''''''''''''
           XTABLAN = XTablaN &  "<table  cellspacing=0 cellpadding=1  style='width:350px;border-style:solid;border-width:1;border-color:#666666;' >"
           XTABLAN = XTablaN &  "<tr HEIGHT=25  >"
           XTABLAN = XTablaN &  "<td " & xsclase & " style='cursor:pointer;' title='Consultar.' " &  sFuncion & "  >" & xtmsgtit & "</TD>"
           XTABLAN = XTablaN &  "</TR>"
           XTABLAN = XTablaN &  "</table>"

''''''''''''''''''
           XTABLAN = XTablaN &  "<table cellspacing=0 cellpadding=0 style='width:350px;border-style:dotted;border-width:1;border-color:#bbbbbb;border-top-style:solid;border-top-width:1;border-top-color:#ffffff;' >"



                         XTMSGSEND = ""
                       '  XTMSGSEND = XTMSGSEND & "<br><B><font color='#aa0000' size=2 >&#0124;</font></B>"
                         XTMSGSEND = XTMSGSEND & "<font color=#888888 style='font-weight:bold;' size=1 >" & tabprocesos (iproc,0,0) & "</font>"

                         XTABLAN = XTablaN &  "<tr >"
                         XTABLAN = XTablaN &  "<td class=UT align='right' colspan=7 STYLE='background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"

                        sAprobo  =  "&nbsp;"

                     If smenu = "PORESTADO"  Then

                        sAprobo = "<a href=coobjwfldet.asp?Tipobjeto=" & tabprocesos(iproc,iprdoc,14) & "&Codobjeto="
                        sAprobo = sAprobo & tabprocesos (iproc,iprdoc,2) & " target=_blank><IMG SRC=Imagenes\ICorSInfoUser.jpg title=" & Chr(34)
                        sAprobo = sAprobo & " Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & Chr(34) & " border=0></a>"

                        sRef     = "coobjwfldet.asp?Tipobjeto=" & tabprocesos(iproc,iprdoc,14) & "&Codobjeto=" & tabprocesos (iproc,iprdoc,2)
                        sImagenP = "&#0232;"
                        sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
                        sinput   = "<a class='ft1n' " & sFuncAnc & " href='" & sRef & "' "
                      '  sinput   = sinput  & "' title="&chr(34)&"Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & tabprocesos (iproc,iprdoc,1) & chr(34) & " target='_blank' >" & sImagenP & "</a>"
                        sinput   = sinput  & "' title="&chr(34)&"Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;" & chr(34) & " target='_blank' >" & sImagenP & "</a>"
                        sAprobo  =  sinput

                     END IF




                         XTABLAN = XTablaN &  "<td class=UT align='center'  STYLE='background-color:#dddddd;' >" & sAprobo & "&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr >"
                             XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
                         for yy=1 to 6
                             XTABLAN = XTablaN &  "<td class=UT width='15%'  >&nbsp;</TD>"
                         next
                             XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                    IF (xtmsgIMG & "") = "1" THEN


                    Strsql = "Select * From objinfoIMG  where codobjeto = 0" & xtmcod & " ORDER BY objSECUEN, OBJID"
                    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                    IF NOT RS1.EOF THEN

                       XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
                       XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

                       XTABLAN = XTablaN &  "<table cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#990000;' >"
                       XTABLAN = XTablaN &  "<tr   >"

                       XSFSTYLE= " STYLE='display:block;' "

                       DO WHILE NOT rs1.eof

                          XTABLAN = XTablaN &  "<td CLASS=UT " & XSFSTYLE & " ><IMG SRC='" & RS1(6) & "' ></TD>"
                          XSFSTYLE= " STYLE='display:none;' "

                        RS1.MOVENEXT
                       LOOP

                         XTABLAN = XTablaN &  "</TR>"
                         XTABLAN = XTablaN &  "</table>"

                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                    END IF






                    ELSEIF (xtmsgIMG & "") = "2" THEN


                         XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
                         XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
                         XTABLAN = XTablaN &  xtmsgVID
                      '   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                    END IF





                         xtmsgCOP = tabprocesos (iproc,iprdoc,4)


                         XTMSGSEND = "<B><font color=#888888 >" & gtfechadorsh(tabprocesos (iproc,iprdoc,5)) & "</font></B>"
                     '    XTMSGSEND = XTMSGSEND & "<B><font color='#aa0000' size=2 >&#0124;</font></B>"
                     '    XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-weight:bold;' >&nbsp;" & xtmdtema & "&nbsp;-&nbsp;" & xtmdstema & "&nbsp;</font>"

                     '    XTMSGSEND = XTMSGSEND & "<br><B><font color='#aa0000' size=2 >&#0124;</font></B>"
                     '    XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-weight:bold;' >&nbsp;" & tabprocesos (iproc,0,0) & "</font>"

                     '    XTABLAN = XTablaN &  "<tr height=25 >"

                       '  XTABLAN = XTablaN &  "<td class=ut VALIGN='TOP' align='center' ><font color='#AA0000' size=2 >&#0149;</font></TD>"
                     '    XTABLAN = XTablaN &  "<td class=uti colspan=7 STYLE='color:#000000;' >"
                     '    XTABLAN = XTablaN &   XTMSGSEND & "</TD>"

						 XTMSGSENDVIGENCIA = "<B><font color=#888888 >" & gtfechadorsh(tabprocesos (iproc,iprdoc,13)) & "</font></B>"
						 XTMSGSENDVIGENCIA = XTMSGSENDVIGENCIA & "<B><font color='#aa0000' size=2 >&#0124;</font></B>"
						 if smenu = "PORESTADO" AND SIDX = "1" THEN
					    	 XTMSGSENDVIGENCIA = XTMSGSENDVIGENCIA & "<font color=#000000 style='font-weight:bold;' >FECHA&nbsp;DE&nbsp;CREACION</font>"

						 ELSE
					    	 XTMSGSENDVIGENCIA = XTMSGSENDVIGENCIA & "<font color=#000000 style='font-weight:bold;' >FECHA&nbsp;DE&nbsp;VIGENCIA</font>"
                         END IF

						 XTABLAN = XTablaN &  "<TR><td class=uti colspan=7 STYLE='color:#000000;' >"
                         XTABLAN = XTablaN &   XTMSGSENDVIGENCIA & "</TD> </TR>"

                         XTABLAN = XTablaN &  "<td class=ut  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTMSGSEND = ""
                     ''''    XTMSGSEND = XTMSGSEND & "<br>"
                         XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-size:8pt;font-weight:normal;' >" & xtmsgCOP & "</font>"

                         XTABLAN = XTablaN &  "<tr  >"
                         XTABLAN = XTablaN &  "<td class=ut VALIGN='TOP' align='center' ><font color='#AA0000' size=2 >&#0149;</font></TD>"
                         XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
                         XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
                         XTABLAN = XTablaN &  "<td class=ut  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"





           XTABLAN = XTablaN & "</table>"
           XTABLAN = XTablaN & "</div >"


           next



            XTABLAN = XTablaN & chr(13)





        next

           sHTML = sHTML  &   XTABLAN


    end if








  '  sHTML = sHTML  &  "</DIV >"







  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""
 '   IF SMENU = "PORREVISION" THEN
 '         sAccion = "COFECREVISION.asp?Menu=" & sMenu & "&Nombre=" & sNombre
 '      ELSE
 '         sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombre
 '   END IF

 'Celda   = ofv.BotonVolver(sAccion)

  'sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

    sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", nlinea, Addqry, "", "no", "", Listar, Celda)
    sHTML = sHTML & "</body></html>"



 Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

'''''''''''''

end if

response.write sHTML


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
