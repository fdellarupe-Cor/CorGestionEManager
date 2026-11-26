<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

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


     If smenu = "PORFECHA" OR smenu = "PORREVISION" Then

      if (SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1) or (SESSION("EMPMUN") = "N") THEN
         addfrom =  " ,objetos a, tipobjeto  "
         sverestado = sverestado & " and csporfechas.CODOBJETO = a.CODOBJETO "
         sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
         sverestado = sverestado & " and a.OBJUN = " & Session("UN") & " "
      ELSE
         addfrom =  " ,OBJMEMPALC ,objetos a, tipobjeto "
         sverestado = sverestado & " and csporfechas.CODOBJETO = a.CODOBJETO "
         sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
         sverestado = sverestado & " and "
         sverestado = sverestado & " ( (a.OBJUN = " & Session("UN") & ") OR "
         sverestado = sverestado & "   ( a.OBJUN = 1 AND "
         sverestado = sverestado & "     csporfechas.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
         sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
         sverestado = sverestado & " ) "
      END IF

         sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "
         sverestado = sverestado & " and (a.APLIC = '" & wkaplic & "' or a.APLIC = 'GLB') "

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



'''''''''''
    If Request.Form.Count > 1 Then
       Robj = "F"
    else
       Robj = "Q"
    end if

    If smenu = "PORREVISION" Then

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

        xsid        = Replace(sID, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID  = "") Then  sID  = "1-1-1950"

        xsid        = Replace(sID2, "/", "")
        xsid        = Replace(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

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
    strsql = strsql & "OBJETOS.OBJUN,OBJPRTY.fecpublic, t.descri,    "
    Strsql = Strsql & "t.descri, t.descri, t.descri, st.descri   "


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
            strsql = strsql & " ((OBJACCION.codobjeto = OBJETOS.codobjeto and OBJACCION.codcomp = objlnk.objid and "
            strsql = strsql & "objlnk.iden2 = objaccion.codaccion) or "
            strsql = strsql & " (objestdet.codobjeto = OBJETOS.codobjeto and "
            strsql = strsql & " objestdet.detimagen = 'J' and "
            strsql = strsql & " objestdet.detimagen = objlnk.objid and "
            strsql = strsql & " objlnk.iden2 = objestdet.objid) ) and "
          '  strsql = strsql & " objlnk.iden2 = CONVERT(nvarchar(50), dbo.OBJESTDET.objid)) ) and "

            strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado = Estados.objid and "
            strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
            strsql = strsql & " and Complementos.descri = '" & sID & "'"



            sTitulo = "Componente: " & sID




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
               strsql = strsql & " and eco.descri = '" & sID2 & "'"

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

               sTitulo = sTitulo & "&nbsp;|&nbsp;Sector: " & sID2

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
                strsql = strsql & "  ((OBJPRTY.instrumen=" & CDbl(sID) & ")"
            Else
                strsql = strsql & "  (OBJPRTY.instrumen=" & CDbl(sID) & " or OBJPRTYADV.instrua=" & CDbl(sID) & ")  "
            End If

            If sID2 <> "" Then
                strsql = strsql & " AND (OBJPRTY.instid ='" & sID2 & "')) OR "
                strsql = strsql & " (OBJPRTYADV.instrua =" & sID & " and OBJPRTYADV.instid ='" & sID2 & "')"
            End If

            strsql = strsql & " ) "

            xinstrum = ""
            Sqlx = "SELECT descri from instrumen where objid=" & CDbl(sID)
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
                    strsql = strsql & "and eco.descri = '" & sID & "'"
                Case Else
                    strsql = strsql & "FROM objcontrol,objproced,OBJETOS, TIPOBJETO, OBJPRTY, Estados, eco, OBJSECTOR " & addfrom
                    strsql = strsql & " where    "
                    strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
                    strsql = strsql & "OBJPRTY.estado = Estados.objid and "
                    strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto and "
                    strsql = strsql & "ECO.eco = OBJSECTOR.codsector and "
                    strsql = strsql & "OBJETOS.codobjeto = OBJSECTOR.codobjeto "
                    strsql = strsql & "and eco.descri = '" & sID & "'"

                    strsql = strsql & " and(( objproced.codobjeto=OBJETOS.codobjeto "
                    strsql = strsql & " and objproced.codsector=OBJSECTOR.OBJID"
                    strsql = strsql & " and objproced.codperfil='" & SPerfil & "')"
                    strsql = strsql & " or (objcontrol.codobjeto=OBJETOS.codobjeto "
                    strsql = strsql & " and objcontrol.codsector=OBJSECTOR.codsector "
                    strsql = strsql & " and objcontrol.codperfil='" & SPerfil & "'))"
            End Select

            sTitulo = "Sector:&nbsp;" & sID
            if  SPerfil <> "nsnc" then sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & SPerfil

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
            cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
            chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
            If sID3 = "Emision" Then
                strsql = strsql & "WHERE   xfecemis between '" & cdesde & "' and '" & chasta & "' "
                sTitulo = "Creacion&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
            ElseIf sID3 = "Vigencia" Then
                strsql = strsql & "WHERE   xfecvigen between '" & cdesde & "' and '" & chasta & "' "
                sTitulo = "vigencia&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
            ElseIf sID3 = "Eliminacion" Then
                strsql = strsql & "WHERE   xfecbaja between '" & cdesde & "' and '" & chasta & "' "
                sTitulo = "Eliminacion&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
            ElseIf sID3 = "Vencimiento" Then
                strsql = strsql & "WHERE  xfecvto between '" & cdesde & "' and '" & chasta & "' "
                sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
            ElseIf sID3 = "Publicacion" Then
                strsql = strsql & "WHERE   xfecpublic between '" & cdesde & "' and '" & chasta & "' "
                sTitulo = "Publicacion&nbsp;&nbsp;Del:&nbsp;" & sID & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
            End If

        Case "PORREVISION"
            strsql = "select * from csporfechas " & addfrom
            cdesde = CDate(sID  + CLng(CDate("01-01-1900")) )
            chasta = CDate(sID2 + CLng(CDate("01-01-1900")) )

            strsql = strsql & "WHERE  xfecvto between '" & sID & "' and '" & sID2 & "' "
            sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & cdesde & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & chasta


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
                        strsql  = strsql & " order by tipobjeto.descri,objdescri, csporfechas.descri"
                Case "T"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by  tipobjeto.descri,csporfechas.descri,objdescri"
                Case "C"
                        sTitOrd = sTitOrd & "Creacion"
                        strsql  = strsql & " order by  tipobjeto.descri,fecemis,objdescri, csporfechas.descri"
                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by  tipobjeto.descri,estdescri,objdescri, csporfechas.descri"
        End Select
    Else
        If Sfiltro <> "TODOS" Then
           sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
           strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
        End If

        Select Case sorden
                Case "N"
                        sTitOrd = sTitOrd & "Nombre"
                        strsql  = strsql & " order by tipobjeto.descri,objetos.descri "
                Case "T"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by tipobjeto.descri,objetos.descri "
                Case "C"
                        sTitOrd = sTitOrd & "Creacion"
                        strsql  = strsql & " order by tipobjeto.descri,objprty.fecemis "
                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by tipobjeto.descri,estdescri,objetos.descri "
        End Select
    End If

 '''response.write "<br><br><br>" & strsql & "<br>" & OFV.strconn2

'''response.write "<br><br><br>" & strsql & "<br>"

    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")



  sHTML      = ofv.FormHeader(session("FormN"))



    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)

    xs = xs & "function Axn(c1,c2,c3) {" & Chr(13)

    xs = xs & " var mydat= 'Consultar.asp?DOCMODAL=CL&Codigo=' + c1 + '&Tipo=' + c2 + '&Perfil=' + c3 + '&Datos=&Volver=" & volver2 & "'; "

    xs = xs & " myWin= open(mydat"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)


    xs = xs & "</script>" & Chr(13)

    sHTML = sHTML & xs


  Fila    = ""
  Celda   = ""
  sHref   = ""

  dim tabprocesos (100,200,20)
  dim tabprocord  (100,5)


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
  TABLAB = TablaB & "<DIV id='DTindex' STYLE='float: left; z-index:1;overflow:NONE;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"


    sInputT = "&nbsp;&nbsp;&nbsp;<B><font color='#0000AA' size=2 >&#0124;</font></B>&nbsp;&nbsp;<font color='#000000' size=1 ><B>" & sTitulo & "</B></font>&nbsp;"

    celda  = ""
    Celda      = Celda & ofv.GenCelda("","vti STYLE='cursor:pointer;color:#990000;font-size:9pt;' ","","","","","","Busqueda&nbsp;por&nbsp;" & sInputT,"si")
  '''  Celda      = Celda & ofv.GenCelda("","UT  ","","70%","","","",sInputT,"si")

    FILA     = ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","","","",Celda,"si")

  TABLAB = TablaB & ofv.GenTabla("","aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ","","98%","","0","","-2","0","","",Fila,"si")

  TABLAB = TablaB & "</DIV>"

           sHTML = sHTML  &   TABLAB


       xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "



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
 '''   response.write strsql & "<br>" & chr(13)
    sVino = "N"

    tipoant = "NO"
    canttipo = 0

    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If rs.EOF Then

       sVino = "N"


    Else



        sVino = "N"

        nlinea = 0
        Do Until rs.EOF
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

                sitem = Capitalize(rs(2) & "")



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


               tabprocesos (iproc,iprdoc,1) = sitem                        'nombre
               tabprocesos (iproc,iprdoc,2) = rs(3)                        'codigo
               tabprocesos (iproc,iprdoc,3) = sPerfil                      'perfil 


               if wspwkcopete = "" then
                  tabprocesos (iproc,iprdoc,4) = "Documento normativo"        'copete
               else
                  tabprocesos (iproc,iprdoc,4) = Capitalize(ltrim(wspwkcopete))      'copete
               end if



               tabprocesos (iproc,iprdoc,5) = rs(6)                        'fec.vigencia


             '  If clng("0" & rs(12)) = 1 and  SESSION("UN") <> 1 and  SESSION("EMPMUN") = "S" Then
             '     tabprocesos (iproc,iprdoc,6) = 1                            'corporativo 0-no 1-si
             '  Else
                  tabprocesos (iproc,iprdoc,6) = 0                            'corporativo 0-no 1-si
             '  End If

               tabprocesos (iproc,iprdoc,7)  = stitem                        'stitem
               tabprocesos (iproc,iprdoc,8)  = rs(22)                        'fec.public
               tabprocesos (iproc,iprdoc,9)  = wspwkimagen                   'objimagen
               tabprocesos (iproc,iprdoc,10) = wspwkurldsc                   'objurldsc
               tabprocesos (iproc,iprdoc,11) = rs(26)                        'tema
               tabprocesos (iproc,iprdoc,12) = rs(27)                        'subtema


            end if

            rs.movenext
            nlinea = nlinea + 1
        Loop





    End If

    if  sVino = "N" then

        Celda   = ""

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "No se encuentran documentos para esta busqueda."
        Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","",sLink,"si")

        Fila    = Fila & ofv.GenRow("","aaa bgcolor='" & xTrColor & "' ","","","25","","",Celda,"si")
        sHTML = sHTML & Fila

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


        for ii=1 to ttabprocesos

               iproc = tabprocord  (ii,0)


          ''     tabprocesos (iproc,0,0)  '' descripcion del tipo de documento


          ''     tabprocesos (iproc,0,1)  '' cantidad de documentos 


            XTABLAN = XTablaN & chr(13)

            XTABLAN = XTablaN & "<div id='DTNOV" & ii & "' style='float: left; display: inline ;  "


           XTABLAN = XTablaN & " width:30%;"

           XTABLAN = XTablaN & "border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' >"


           XTABLAN = XTablaN &  "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
           XTABLAN = XTablaN &  "<tr HEIGHT=25  >"
           XTABLAN = XTablaN &  "<td " & xsclase & "   >" & Capitalize(tabprocesos (iproc,0,0)) & "</TD>"
           XTABLAN = XTablaN &  "</TR>"
           XTABLAN = XTablaN &  "</table>"


           XTABLAN = XTablaN & "<div id='DTNOVINT" & ii & "' style='"

           XTABLAN = XTablaN & "border-style:solid;border-width:1;border-color:#eeeeee;' >"
 
           XTABLAN = XTablaN &  "<table width=99% cellspacing=0 cellpadding=0  >"


           for iprdoc=1 to tabprocesos (iproc,0,1)


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




                       '  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                       '  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                       '  XTABLAN = XTablaN &  "</TR>"



                         xtmsgtit = tabprocesos (iproc,iprdoc,1)


                         if (dFechai - tabprocesos (iproc,iprdoc,8)) < 31 then


                            xtmsgtit = "<B><font color='#ffffff' STYLE='background-color:#ff0000;' size=1 >&nbsp;nuevo&nbsp;</font></B>&nbsp;&nbsp;" & xtmsgtit

                         end if


                         XTABLAN = XTablaN &  "<tr HEIGHT='32' bgcolor='#f0f0f0' >"
                         XTABLAN = XTablaN &  "<td class=Vti colspan=8 STYLE='cursor:pointer;color:#990000;font-size:9pt;"
                         XTABLAN = XTablaN &  "border-top-style:solid;border-top-width:3;border-top-color:#ffffff;' "
                         XTABLAN = XTablaN &  sFuncion & " title='Consultar.' >" & xtmsgtit & "</TD>"
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
                         XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                         XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-weight:bold;' >&nbsp;" & xtmdtema & "&nbsp;-&nbsp;" & xtmdstema & "</font>"
                         XTMSGSEND = XTMSGSEND & "<br>"
                         XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-size:8pt;font-weight:normal;' >" & xtmsgCOP & "</font>"

                         XTABLAN = XTablaN &  "<tr  >"
                         XTABLAN = XTablaN &  "<td class=ut VALIGN='TOP' align='center' ><font color='#AA0000' size=2 >&#0149;</font></TD>"
                         XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
                         XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
                         XTABLAN = XTablaN &  "<td class=ut  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;' >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"



           next


           XTABLAN = XTablaN & "</table>"
           XTABLAN = XTablaN & "</div >"



           XTABLAN = XTablaN & "</div >"

            XTABLAN = XTablaN & chr(13)





        next

           sHTML = sHTML  &   XTABLAN


    end if 








  '  sHTML = sHTML  &  "</DIV >"




    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)


  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


    sHTML = sHTML & "</body></html>"





'''''''''''''

end if

response.write sHTML





%>
