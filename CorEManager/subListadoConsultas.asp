<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else



    Dim bconsultar()

    
    Generar = ""
    
    Addqry = ""
    Response.expires = 0

    Server.scripttimeout = 600
    Response.buffer = True
    smenu = Request.querystring("Menu")
    
    sorden = Request.querystring("sorden")
    
    If Len(sorden) = 0 Then
      sorden = "N"
    End If
    
    Sfiltro = Request.querystring("sfiltro")
    
    If Len(Sfiltro) = 0 Then
      Sfiltro = "TODOS"
    End If

    Sfiltro = Replace(Sfiltro, "_", " ")
    CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
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

    If smenu = "PORFECHA" Then
        sverestado = " and objnivel >= " & Session("SegNivel") & " And Estcod In (9,"
    Else
        sverestado = " and Objetos.objnivel >= " & Session("SegNivel") & " And Estados.objid In (9,"
    End If

    For i = 1 To 4
        If bconsultar(i) Then sverestado = sverestado & i & ","
    Next
    sverestado = Left(sverestado, Len(sverestado) - 1) & ") "

    If smenu = "PORTEMA" Then
        bLista = "on"
        besquema = "on"
        sID = Request.querystring("ID")
        
        sMenuAdd = "*ID=" & Request.querystring("ID")
        sMenuMas = "&smenu=" & smenu & "&lista=" & Request.Form("lista") & "&esquema="
        sMenuMas = sMenuMas & Request.Form("esquema")
    ElseIf smenu = "PORFECHA" Then

        If Request.Form.Count > 0 Then
        
           sID = Request.Form("Desde")
           sID2 = Request.Form("Hasta")
           sID3 = Request.Form("Fecha")
           Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
        Else
           sID = Request.querystring("Desde")
           sID2 = Request.querystring("Hasta")
           sID3 = Request.querystring("Fecha")
           Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
        End If
        
        
        
        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
        sMenuMas = sMenuMas & sID2 & "&Fecha=" & sID3

        
        xsid = OFV.convertircar(sID, "/", "")
        xsid = OFV.convertircar(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID = "") Then sID = "1-1-1950"
        xsid = OFV.convertircar(sID2, "/", "")
        xsid = OFV.convertircar(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"
        
    ElseIf smenu = "PORINSTRUMENTO" Then

        If Request.Form.Count > 0 Then
        
          sID = Request.Form("Lista")
          sID2 = Request.Form("buscar")
          Addqry = "&Lista=" & sID & "&buscar=" & sID2
          sID2 = OFV.convertircar(sID2, "_", " ")
        End If
        ElseIf smenu = "PORTEXTO" Then

        If Request.Form.Count > 0 Then
        
         
          sID = Replace(Request.Form("buscar"), "'", "")
          sID = Replace(sID, Chr(34), "")
          Addqry = "&buscar=" & sID
          sID = OFV.convertircar(sID, "_", " ")

        Else
      
            sID = Request.querystring("buscar")
            Addqry = "&Lista=" & sID & "&buscar=" & sID
            sID = OFV.convertircar(sID, "_", " ")

        End If


        
        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&Instrumento=" & sID & "&buscar="
        sMenuMas = sMenuMas & sID2
   ElseIf smenu = "PORIDENTIFICACION" Then

        If Request.Form.Count > 0 Then
        
          sID = Request.Form("buscar")
          Addqry = "&buscar=" & sID

        Else
           sID = Request.querystring("buscar")
           Addqry = "&buscar=" & sID
        End If


        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&IDENTIFICACION=" & sID
    Else

        If Request.Form.Count > 0 Then
        SPerfil = OFV.convertircar(Trim(Request.Form("perfil")), "_", " ")
        sID = Request.Form("lista")
        sID2 = Request.Form("lista")
        sID3 = Request.Form("buscar")
        Addqry = "&Lista=" & sID2 & "&buscar=" & sID3

        Else
           SPerfil = Trim(Request.querystring("perfil"))
           sID = Request.querystring("Lista")
           sID2 = Request.querystring("Lista")
           sID3 = Request.querystring("buscar")
           Addqry = "&Lista=" & sID2 & "&buscar=" & sID3

        End If


        
        sID = OFV.convertircar(sID, "_", " ")
        If sID = "" Then sID = sID3
        
        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
        sMenuMas = sMenuMas & sID3
    End If
    Addqry = "&Menu=" & smenu & Addqry
    starget = "subListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry  &"&perfil="&replace(sperfil," ","_")
    sVolver = "subListadoConsultas.asp?Menu=" & smenu & sMenuAdd &"&perfil="&replace(sperfil," ","_")
    Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")
    Set cn = OFV.conectar(OFV.strconn2)

    strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
    strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, "
    strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, "
    strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
    strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover "

    Select Case smenu
    Case "PORCOMPONENTE"
'        strsql = strsql & "FROM (((OBJACCION INNER JOIN COMPLEMENTOS ON "
'        strsql = strsql & "OBJACCION.codcomp = COMPLEMENTOS.codcomp) INNER JOIN OBJETOS ON "
'        strsql = strsql & "OBJACCION.codobjeto = OBJETOS.codobjeto) INNER JOIN TIPOBJETO ON "
'        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto) INNER JOIN (OBJPRTY "
'        strsql = strsql & "INNER JOIN Estados ON OBJPRTY.estado = Estados.objid) ON "
'        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

        strsql = strsql & "FROM OBJACCION, COMPLEMENTOS, OBJETOS, TIPOBJETO, OBJPRTY, Estados  "
        strsql = strsql & " where "
        strsql = strsql & "OBJACCION.codcomp = COMPLEMENTOS.codcomp and "
        strsql = strsql & "OBJACCION.codobjeto = OBJETOS.codobjeto and "
        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
        strsql = strsql & "OBJPRTY.estado = Estados.objid and "
        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
        strsql = strsql & " and Complementos.descri = '" & sID & "'"
        sTitulo = "Componente: " & sID
    Case "PORESTADO"
'        strsql = strsql & "FROM (OBJETOS INNER JOIN TIPOBJETO ON "
'        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto) INNER JOIN (OBJPRTY "
'        strsql = strsql & "INNER JOIN Estados ON OBJPRTY.estado = Estados.objid) ON "
'        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

        strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "
        strsql = strsql & " where "
        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
        strsql = strsql & "OBJPRTY.estado = Estados.objid and "
        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
        strsql = strsql & "and OBJPRTY.estado = " & sID
        sTitulo = "Estado: " & sID
        sqle = "Select * from estados where objid =  " & sID
        Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
        If Not rsx.EOF Then sTitulo = "Estado: " & rsx(1)
        
    Case "PORTEXTO"
        sTitulo = "Texto: " & sID
       
        strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "
        strsql = strsql & " where "
        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
        strsql = strsql & "OBJPRTY.estado = Estados.objid and "
        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
        strsql = strsql & "and (OBJETOS.descriabr like '%" & sID & "%' "
        strsql = strsql & " or OBJETOS.descri  like '%" & sID & "%') "
        'Response.write strsql & "<br>"
        sTitulo = "Texto: " & sID
        'sqle = "Select * from estados where objid =  " & sID
       ' Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
       ' If Not rsx.EOF Then sTitulo = "texto: " & rsx(1)
        
    Case "PORINSTRUMENTO"
'        strsql = strsql & "FROM (OBJETOS INNER JOIN TIPOBJETO ON "
'        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto) INNER JOIN (OBJPRTY "
'        strsql = strsql & "INNER JOIN Estados ON OBJPRTY.estado = Estados.objid) ON "
'        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

        strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "
        strsql = strsql & " where "
        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
        strsql = strsql & "OBJPRTY.estado = Estados.objid and "
        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
        strsql = strsql & "and OBJPRTY.instrumen = " & CDbl(sID)
        If sID2 <> "" Then strsql = strsql & " AND OBJPRTY.instid = '" & sID2 & "'"
        xinstrum = ""
        Sqlx = "SELECT descri from instrumen where objid=" & CDbl(sID)
        Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
        If Not rsx.EOF Then xinstrum = rsx(0)
        Call OFV.cerrarconsulta(rsx)
        sTitulo = "INSTRUMENTO " & xinstrum & " " & sID2
    Case "PORIDENTIFICACION"
'        strsql = strsql & "FROM (OBJETOS INNER JOIN TIPOBJETO ON "
'        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto) INNER JOIN (OBJPRTY "
'        strsql = strsql & "INNER JOIN Estados ON OBJPRTY.estado = Estados.objid) ON "
'        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

        strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "
        strsql = strsql & " where "
        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
        strsql = strsql & "OBJPRTY.estado = Estados.objid and "
        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
        strsql = strsql & "and OBJETOS.Codiden = '" & sID & "' "
        sTitulo = "IDENTIFICACION " & sID
    Case "PORSECTOR"
'        strsql = strsql & "FROM ((OBJETOS INNER JOIN TIPOBJETO ON "
'        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto) INNER JOIN (OBJPRTY "
'        strsql = strsql & "INNER JOIN Estados ON OBJPRTY.estado = Estados.objid) ON "
'        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto) INNER JOIN (ECO "
'        strsql = strsql & "INNER JOIN OBJSECTOR ON ECO.eco = OBJSECTOR.codsector) ON "
'        strsql = strsql & "OBJETOS.codobjeto = OBJSECTOR.codobjeto "
     Select Case SPerfil
     Case "nsnc"
     
        strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, eco, OBJSECTOR "
        strsql = strsql & " where "
        strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
        strsql = strsql & "OBJPRTY.estado = Estados.objid and "
        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto and "
        strsql = strsql & "ECO.eco = OBJSECTOR.codsector and "
        strsql = strsql & "OBJETOS.codobjeto = OBJSECTOR.codobjeto "
        strsql = strsql & "and eco.descri = '" & sID & "'"
      Case Else
        strsql = strsql & "FROM objcontrol,objproced,OBJETOS, TIPOBJETO, OBJPRTY, Estados, eco, OBJSECTOR "
        strsql = strsql & " where "
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
        sTitulo = "Sector: " & sID & " Perfil='" & SPerfil & "'"
    Case "PORTEMA"
        sql = "Select descri From Objetos Where codobjeto = " & sID
        Set rsObj = OFV.crearconsultaEx(sql, cn, 1, "")
        If Not rsObj.EOF Then sObj = rsObj(0)
        Call OFV.cerrarconsulta(rsObj)
'        strsql = strsql & "FROM (OBJETOS INNER JOIN TIPOBJETO ON "
'        strsql = strsql & "OBJETOS.tipobjeto=TIPOBJETO.tipobjeto) INNER JOIN (OBJPRTY "
'        strsql = strsql & "INNER JOIN Estados ON OBJPRTY.estado = Estados.objid) ON "
'        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

        strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "
        strsql = strsql & " where "
        strsql = strsql & "OBJETOS.tipobjeto=TIPOBJETO.tipobjeto and "
        strsql = strsql & "OBJPRTY.estado = Estados.objid and "
        strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
        strsql = strsql & "and Objetos.codobjeto = " & sID
        sTitulo = "Tema: " & sObj
    Case "PORFECHA"
        strsql = "select * from csporfechas "
        cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
        chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
        If sID3 = "Emision" Then
            strsql = strsql & "WHERE xfecemis between '" & cdesde & "' and '" & chasta & "' "
            sTitulo = "Fecha de emisi&0acuten: Desde: " & sID & " Hasta: " & sID2
        ElseIf sID3 = "Vigencia" Then
            strsql = strsql & "WHERE xfecvigen between '" & cdesde & "' and '" & chasta & "' "
            sTitulo = "Fecha de vigencia: Desde: " & sID & " Hasta: " & sID2
        ElseIf sID3 = "Eliminacion" Then
            strsql = strsql & "WHERE xfecbaja between '" & cdesde & "' and '" & chasta & "' "
            sTitulo = "Fecha de eliminaci&0acuten: Desde: " & sID & " Hasta: " & sID2
        ElseIf sID3 = "Vencimiento" Then
            strsql = strsql & "WHERE xfecvto between '" & cdesde & "' and '" & chasta & "' "
            sTitulo = "Fecha de Vencimiento: Desde: " & sID & " Hasta: " & sID2
        End If
    End Select

    strsql = strsql & sverestado
    
    If smenu = "PORFECHA" Then
       ' strsql = strsql & " order by objdescri, descri"
    If Sfiltro <> "TODOS" Then
     strsql = strsql & " and descri='" & Sfiltro & "' "
    End If
    
    Select Case sorden
    Case "N"
        strsql = strsql & " order by objdescri, descri"
    Case "T"
        strsql = strsql & " order by  descri,objdescri"
    Case "C"
        strsql = strsql & " order by  fecemis,objdescri, descri"
    Case "E"
        strsql = strsql & " order by  estdescri,objdescri, descri"
    End Select
    Else
    If Sfiltro <> "TODOS" Then
     strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
    End If
    
    Select Case sorden
    Case "N"
        strsql = strsql & " order by objetos.descri,tipobjeto.descri "
    Case "T"
        strsql = strsql & " order by tipobjeto.descri,objetos.descri "
    Case "C"
        strsql = strsql & " order by objprty.fecemis,objetos.descri,tipobjeto.descri "
    Case "E"
        strsql = strsql & " order by estdescri,objetos.descri,tipobjeto.descri "
    End Select
    End If

   ' response.write "sql: " & strsql & "<br>"
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

    sHTML = OFV.FormHeader(sTitulo)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs






    Celda = OFV.GenCelda("", "ts colspan=2", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
    Celda = Celda & OFV.GenCelda("", "ts", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=T target=_self ><B><u>&nbsp;TIPO&nbsp;</u></B></A>", "si")
    If smenu = "PORESTADO" Then
        If sID = 4 Then
            Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Eliminacion", "si")
        Else
            Celda = Celda & OFV.GenCelda("", "ts", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CREACION&nbsp;</u></B></A>", "si")
        End If
    Else
        Celda = Celda & OFV.GenCelda("", "ts", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CREACION&nbsp;</u></B></A>", "si")
    End If
    If smenu = "PORFECHA" Then
        If sID3 = "Eliminacion" Then
            Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Eliminacion", "si")
        ElseIf sID3 = "Vencimiento" Then
            Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Vencimiento", "si")
        Else
            Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Vigencia", "si")
        End If
    Else
        Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Vigencia", "si")
    End If
    If smenu <> "PORESTADO" Then
        Celda = Celda & OFV.GenCelda("", "ts", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=E target=_self ><B><u>&nbsp;ESTADO&nbsp;</u></B></A>", "si")
    Else
        Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Autor", "si")
        If sID = 3 Or sID = 4 Then
            Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Aprobo", "si")
        End If
        If sID = 4 Then
            Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Elimino", "si")
        End If
    End If
    Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    Response.write sHTML
    sHTML = ""
    Response.flush

    nLinea = 0
   ' response.write 
    fi = CLng(CDate("01-01-1900"))
     PosReg = 0
    'Do Until rs.EOF
     On Error Resume Next
     If Not rs.EOF Then rs.Move (CantRegMover)
     Do While (Not rs.EOF) And (PosReg < CantRegMostrar)
        dEmision = "SD"
        dVigencia = "SD"
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
        If dEmi > 0 Then dEmision = "SD"
        If smenu = "PORFECHA" Then
            If sID3 = "Eliminacion" Then
                If Not IsNull(rs(7)) Then dVigencia = CDate(CLng(rs(7)) + fi)
            ElseIf sID3 = "Vencimiento" Then
                If Not IsNull(rs(16)) Then dVigencia = CDate(CLng(rs(16)) + fi)
            Else
                If Not IsNull(rs(6)) Then dVigencia = CDate(CLng(rs(6)) + fi)
            End If
        Else
            If Not IsNull(rs(6)) Then dVigencia = CDate(CLng(rs(6)) + fi)
        End If

'        Else
'            If Not IsNull(rs(6)) Then dVigencia = rs(6)
'        End If

        dVign = InStr(dVigencia, "1900")
        If dVign > 0 Then dVigencia = "SD"
        If smenu <> "PORESTADO" Then
            If Not IsNull(rs(8)) Then
                If (Not CDate(CLng(rs(16)) + fi) > Date) And (CDate(CLng(rs(16)) + fi) <> CDate("01-01-1900")) Then
            sEstado = "VENCIDA"
                Else
                    sEstado = rs(8)
                End If
            End If
        Else
            sAutor = "&nbsp;"
            sAprobo = "&nbsp;"
            sElimino = "&nbsp;"
            If Not IsNull(rs(12)) Then sAutor = rs(12)
            If Not IsNull(rs(13)) Then sAprobo = rs(13)
            If Not IsNull(rs(14)) Then sElimino = rs(14)
        End If
        If Not IsNull(rs(9)) Then nIden = rs(9)
        If Not IsNull(rs(10)) Then sobjeto = rs(10)
        If smenu <> "PORFECHA" Then
            If Not IsNull(rs(19)) Then sobjeto = sobjeto & " - " & rs(19)
            If Not IsNull(rs(20)) Then sobjeto = sobjeto & " " & rs(20)
        End If
        If Session("USRES") <> "B" Then
           sImagen = OFV.GenImage("", "", "", "", sArchivo, "0", "")
        Else
           sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"
        End If

        sActcon = "Consultar.asp?Codigo=" & rs("codobjeto") & "&Tipo=" & rs("tipobjeto")
        sActcon = sActcon & "&Perfil=" & rs("perfil") & "&Datos=" & OFV.convertircar(sMenuMas, "&", "*")
        sActcon = sActcon & "&Volver=" & sVolver
        sHref = "&nbsp;"
        sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)
       '  response.write sactcon
        sLink = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen, "si")

        Celda = OFV.GenCelda("", "", "", "", "", "", "", sLink, "si")

        sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, sobjeto, "si")

        Celda = Celda & OFV.GenCelda("", "abc style={font-size:10;}", "", "", "", "", "", sLink, "si")
        Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=" & sorden & "&sfiltro=" & Replace(stipo, " ", "_") & " target=_self ><font face=verdana size=1 color=#336699><b>" & stipo & "</b></font></A>", "si")
        Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", dEmision, "si")
        Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", dVigencia, "si")

        If smenu <> "PORESTADO" Then
            If sEstado = "VENCIDA" Then
                sInput = "<font color=#880000>" & sEstado & "</font>"
            Else
                sInput = sEstado
            End If
            Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", sInput, "si")
        Else
            Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", sAutor, "si")
            If sID = 3 Or sID = 4 Then
                sAprobo = "<a href=coobjwfldet.asp?Tipobjeto=" & rs("tipobjeto") & "&Codobjeto="
                sAprobo = sAprobo & rs("codobjeto") & " target=_blank><IMG SRC=Imagenes\ICorSInfoUser.jpg title=" & Chr(34)
                sAprobo = sAprobo & " Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & Chr(34) & " border=0></a>"
    
                Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", sAprobo, "si")
            End If
            If sID = 4 Then
                Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", sElimino, "si")
            End If
        End If
        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
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
    sHTML = sHTML & OFV.GenTabla("", "", "", "100% rules=rows style={border-style:solid;border-width:1;border-color:#336699;} ", "", "0", "", "-1", "0", "", "", Fila, "si") & "</body></html>"
    
     
     sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry  & "&perfil="&replace(sperfil," ","_"), "", "no", "", "", "")
    if smenu="PORSECTOR" then
     session("sperfil")=replace(sperfil," ","_")
    else
     session("sperfil")=""
    end if 
    
    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

''     Generar = sHTML
    





end if
'response.write request.form("lista")
r'esponse.write addqry & "&perfil="&sperfil
response.write sHTML
response.flush

%>