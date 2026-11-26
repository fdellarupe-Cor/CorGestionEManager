
Function Generar()


    Generar = ""

    nCodigo = Request.querystring("Codigo")
    sTipo = Request.querystring("Tipo")
    sPerfil = Request.querystring("Perfil")
    bCopiar = Request.querystring("Copiar")
    sOpcion = Request.querystring("Opcion")
    sUsuario = Session("Usuario")
    sTipobjeto = sOpcion
    volver = ""

'''''''''''''
   If bCopiar = "SI" Then
        Server.scripttimeout = 600
        Response.write OFV.FormBody
        Response.write "<p><font face=arial-narrow size=1 color=#336699>  Log de la operacion </p>"
        ntime0 = Minute(Now) & ":" & Second(Now)
        Response.write "tiempo proc inicio: " & ntime0
        xCodigo = nCodigo
        xtipobjeto = sTipobjeto

      sCodiden = UCase(Request.Form("codiden"))
        sDescri = UCase(Request.Form("descri"))
        sEsquema = Request.Form("Esquema")
        i = InStr(sEsquema, "_")
        nEsquema = Left(sEsquema, i - 1)
        sEsquema = Mid(sEsquema, i + 1)
        'Objetos
        Set cn = OFV.conectar(OFV.strconn2)
        nFecini = 0
        If nFecini < CLng(CDate("01-01-1900")) Then
            nFecini = CLng(CDate("01-01-1900"))
        Else
            nFecini = 0
        End If

        dFechaNula = CLng(CDate("01-01-1900")) - nFecini

      strsql = "Select objnivel From Objetos Where codobjeto = " & nCodigo
        Set rsNew = OFV.crearconsulta2(strsql, cn)
      If Not rsNew.EOF Then
            nOBJNIV = rsNew(0)
        Else
            nOBJNIV = 9
        End If
        Call OFV.cerrarconsulta(rsNew)

      wsval1 = sDescri
        strsql = "Select descriabr From Objetos Where codobjeto = " & nCodigo
        Set rsNew = OFV.crearconsulta2(strsql, cn)
        If Not rsNew.EOF Then
           If IsNull(rsNew(0)) Then
              wsval1 = sDescri
           Else
              wsval1 = rsNew(0)
           End If
        End If
        Call OFV.cerrarconsulta(rsNew)


        strsql = "Insert Into Objetos (codiden,tipobjeto,descri,objrem,objrdo,objnivel,descriabr) "
        strsql = strsql & "Values ('" & sCodiden & "','" & sTipobjeto & "','" & Session.sessionid & "',"
        strsql = strsql & nCodigo & ",0," & nOBJNIV & ",'" & wsval1 & "')"
        Call OFV.crearconsultaEx(strsql, cn, 1, volver)

        strsql = "Select codobjeto From Objetos Where descri = '" & Session.sessionid & "'"
        Set rsNew = OFV.crearconsulta2(strsql, cn)
        If Not rsNew.EOF Then
            rsNew.movelast
            nCodigo = rsNew(0)
            Response.write "nCodigo=" & nCodigo & "<br>"
        End If
        Call OFV.cerrarconsulta(rsNew)

        strsql = "Update Objetos set descri = '" & sDescri & "' Where descri = '" & Session.sessionid & "'"
        Call OFV.crearconsultaEx(strsql, cn, 1, volver)
        Call OFV.cerrarconn(cn)

    'Relaciones
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Insert into Relaobj (Codobjmaster,Tipobjmaster,Codojasoc,Tipobjasoc) "
        strsql = strsql & "Values (" & nEsquema & ",'" & sEsquema & "'," & nCodigo & ",'"
        strsql = strsql & sTipobjeto & "')"
        Call OFV.crearconsultaEx(strsql, cn, 1, volver)
        Call OFV.cerrarconn(cn)

        'Propietario
        Set cn = OFV.conectar(OFV.strconn2)

        strsql = "Insert into OBJSDK (codobjeto, tipobjeto, segct,segmo,segco,segce,segown "
        strsql = strsql & ") Values (" & nCodigo & ",'" & sTipobjeto & "',0,0,0,0,'"
        strsql = strsql & Session("Usuario") & "')"

        Call OFV.crearconsultaEx(strsql, cn, 1, volver)
        Call OFV.cerrarconn(cn)

    'Propiedades
        Nfechoy = CLng(Date) - nFecini
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objprty Where codobjeto = " & xCodigo
        Set rsPrty = OFV.crearconsultaEx(strsql, cn, 1, volver)
        strsql = "Insert Into Objprty (codobjeto,tipobjeto,origen,instrumen,instid,estado,"
        strsql = strsql & "autor,fecemis,fecvigen,fecbaja,fecvto,version,nrover) Values (" & nCodigo & ",'"
        strsql = strsql & sTipobjeto & "'," & rsPrty(3) & "," & rsPrty(4) & ",'" & rsPrty(5)
        strsql = strsql & "',1,'" & sUsuario & "','" & Nfechoy & "','" & dFechaNula & "','"
        strsql = strsql & dFechaNula & "','" & dFechaNula & "','" & rsPrty(14) & "','" & rsPrty(15) & "')"
        Call OFV.crearconsultaEx(strsql, cn, 1, volver)
        Call OFV.cerrarconsulta(rsPrty)
        Call OFV.cerrarconn(cn)
        Response.write "<p> Propiedades 1</p>"
' '''''''
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select Objperfil.* From Objperfil, Tipobjeto where "
        strsql = strsql & "Objperfil.tipoperfil = Tipobjeto.perfil "
        strsql = strsql & "and Tipobjeto.tipobjeto = '" & xtipobjeto & "'"
        Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)

        bSector = "NO"
        bEvento = "NO"
        bProcedimiento = "NO"
        bControl = "NO"
        bEsquema = "NO"
        bTablas = "NO"
        bDocum = "NO"
        bEditor = "NO"
        bDescri = "NO"
        bMain = "NO"



      If rsPerfil("objsector") Then
            bSector = "SI"
            'Sectores
        End If
        If rsPerfil("objeven") Then
            bEvento = "SI"
            'Eventos
        End If
        If rsPerfil("objproced") Then
            bProcedimiento = "SI"
            'Procedimientos
        End If
        If rsPerfil("objcontrol") Then
            bControl = "SI"
            'Controles
        End If
        If rsPerfil("objesquema") Then
            bEsquema = "SI"
            'Esquemas
        End If
        If rsPerfil("objtabla") Then
            bTablas = "SI"
            'Tablas
        End If
        If rsPerfil("objhtm") Then
            bDocum = "SI"
            'Documentos
        End If
        If rsPerfil("objobj1") Then
            bEditor = "SI"
            'Editor estructurado
        End If
        If rsPerfil("objaccion") Then
            bDescri = "SI"
            'Descripcion
        End If

        If rsPerfil("objmain") Then
            bMain = "SI"
            'Consideraciones Generales
        End If


        Call OFV.cerrarconsulta(rsPerfil)
        Call OFV.cerrarconn(cn)



        Set cn = OFV.conectar(OFV.strconn2)
  strsql = "Select * From Objperfildoc Where codobjeto =" & xCodigo
'  response.write "sql " & strsql
  Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)
  If rsPerfil.EOF Then
     perfilsp = "N"
'  response.write "no "
  Else
'  response.write "si "
        bSector = "NO"
        bEvento = "NO"
        bProcedimiento = "NO"
        bControl = "NO"
        bEsquema = "NO"
        bTablas = "NO"
        bDocum = "NO"
        bEditor = "NO"
        bDescri = "NO"
        bMain = "NO"



      If rsPerfil("objsector") Then
            bSector = "SI"
            'Sectores
        End If
        If rsPerfil("objeven") Then
            bEvento = "SI"
            'Eventos
        End If
        If rsPerfil("objproced") Then
            bProcedimiento = "SI"
            'Procedimientos
        End If
        If rsPerfil("objcontrol") Then
            bControl = "SI"
            'Controles
        End If
        If rsPerfil("objesquema") Then
            bEsquema = "SI"
            'Esquemas
        End If
        If rsPerfil("objtabla") Then
            bTablas = "SI"
            'Tablas
        End If
        If rsPerfil("objhtm") Then
            bDocum = "SI"
            'Documentos
        End If
        If rsPerfil("objobj1") Then
            bEditor = "SI"
            'Editor estructurado
        End If
        If rsPerfil("objaccion") Then
            bDescri = "SI"
            'Descripcion
        End If

        If rsPerfil("objmain") Then
            bMain = "SI"
            'Consideraciones Generales
        End If

        strsql = "Insert into Objperfildoc (codobjeto,tipobjeto"
        strsql2 = ") values (" & nCodigo & ",'" & sTipobjeto & "'"

    For ii = 0 To (rsPerfil.fields.Count - 4)
      strsql = strsql & "," & rsPerfil(ii).Name
      s = OFV.QueEsx(rsPerfil(ii).Type)
      If rsPerfil(ii).Type > 100 And rsPerfil(ii).Type < 300 Then s = "'"

      sInput = rsPerfil(ii)

      If UCase(sInput) = "VERDADERO" Then sInput = "1"
      If UCase(sInput) = "FALSO" Then sInput = "0"


      If rsPerfil(ii).Type = 131 Then sInput = Replace(rsPerfil(ii), ",", ".")


      strsql2 = strsql2 & "," & s & sInput & s

    Next

    strsql = strsql & strsql2 & ") "

 ' response.write "sql " & strsql & "<br>"
    Call OFV.crearconsultaEx(strsql, cn, 1, volver)



  End If

        Call OFV.cerrarconsulta(rsPerfil)
        Call OFV.cerrarconn(cn)


'''''''''''



    'Descripcion
        If bDescri = "SI" Then
            Set cn = OFV.conectar(OFV.strconn2)
            strsql = "Select * From Objdscabr Where codobjeto = " & xCodigo
            Set rsEvento = OFV.crearconsultaEx(strsql, cn, 1, volver)
            Do Until rsEvento.EOF
                strsql = "Insert Into Objdscabr (codobjeto,tipobjeto,descri,titulo,muestra) Values ("
                strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & rsEvento(3) & "','"
                strsql = strsql & rsEvento(4) & "','" & rsEvento(5) & "')"
                Call OFV.crearconsultaEx(strsql, cn, 1, volver)
                rsEvento.movenext
            Loop
            Call OFV.cerrarconsulta(rsEvento)
            Call OFV.cerrarconn(cn)
            Response.write "<p><font face=arial-narrow size=1 color=#336699>  Eventos: " & EY & " </p>"
        End If


    'Sectores
        If bSector = "SI" Then
            Set cn = OFV.conectar(OFV.strconn2)
            strsql = "Select * From Objsector Where codobjeto = " & xCodigo & " order by objid"
            Set rsSector = OFV.crearconsultaEx(strsql, cn, 1, volver)
            SY = 0
            si = 0
            Do Until rsSector.EOF
                ReDim Preserve asector(si)
                asector(si) = rsSector(0)
                si = si + 1
                SY = SY + 1
                strsql = "Insert Into Objsector (codobjeto,tipobjeto,calificador,codsector) "
                strsql = strsql & "Values (" & nCodigo & ",'" & sTipobjeto & "','" & rsSector(3)
                strsql = strsql & "'," & rsSector(4) & ")"
                Call OFV.crearconsultaEx(strsql, cn, 1, volver)
                rsSector.movenext
            Loop
            Call OFV.cerrarconsulta(rsSector)
            strsql = "Select * From Objsector Where codobjeto = " & nCodigo & " order by objid"
            Set rsSector = OFV.crearconsultaEx(strsql, cn, 1, volver)
            si = 0
            Do Until rsSector.EOF
                ReDim Preserve nsector(si)
                nsector(si) = rsSector(0)
                si = si + 1
                rsSector.movenext
            Loop
            Call OFV.cerrarconsulta(rsSector)
            Call OFV.cerrarconn(cn)
            Response.write "<p><font face=arial-narrow size=1 color=#336699>  Sectores: " & SY & " </p>"
        End If

    'Eventos
        If bEvento = "SI" Then
            Set cn = OFV.conectar(OFV.strconn2)
            strsql = "Select * From Objeven Where codobjeto = " & xCodigo & " order by objid"
            Set rsEvento = OFV.crearconsultaEx(strsql, cn, 1, volver)
            EY = 0
            ei = 0
            Do Until rsEvento.EOF
                EY = EY + 1
                ReDim Preserve aevento(ei)
                aevento(ei) = rsEvento(0)
                ei = ei + 1
                strsql = "Insert Into Objeven (codobjeto,tipobjeto,descri,secuen) Values ("
                strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & rsEvento(3) & "',"
                strsql = strsql & rsEvento(4) & ")"
                Call OFV.crearconsultaEx(strsql, cn, 1, volver)
                rsEvento.movenext
            Loop
            Call OFV.cerrarconsulta(rsEvento)
            strsql = "Select * From Objeven Where codobjeto = " & nCodigo & " order by objid"
            Set rsEvento = OFV.crearconsultaEx(strsql, cn, 1, volver)
            ei = 0
            Do Until rsEvento.EOF
                ReDim Preserve nevento(ei)
                nevento(ei) = rsEvento(0)
                ei = ei + 1
                rsEvento.movenext
            Loop
            Call OFV.cerrarconsulta(rsEvento)
            Call OFV.cerrarconn(cn)
            Response.write "<p><font face=arial-narrow size=1 color=#336699>  Eventos: " & EY & " </p>"
        End If

    'Documentos
        If bDocum = "SI" Then
            Set cn = OFV.conectar(OFV.strconn2)
            strsql = "Select * From Objhtm Where codobjeto = " & xCodigo & " order by objid"
            Set rshtm = OFV.crearconsultaEx(strsql, cn, 1, volver)
            EY = 0
            ei = 0
        Do Until rshtm.EOF
           EY = EY + 1
           ei = ei + 1
          strsql = "Insert Into Objhtm (codobjeto,tipobjeto,secuen,descri,archivo) Values ("
         strsql = strsql & nCodigo & ",'" & sTipobjeto & "'," & rshtm(3) & ",'"
          strsql = strsql & rshtm(4) & "','" & rshtm(6) & "')"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          rshtm.movenext
        Loop
      Call OFV.cerrarconsulta(rshtm)
      Call OFV.cerrarconn(cn)
      Response.write "<p><font face=arial-narrow size=1 color=#336699>  Documentos: " & EY & " </p>"
      End If

  'Tablas
      If bTablas = "SI" Then
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objtabdef Where codobjeto = " & xCodigo & " order by objid"
        Set rstablas = OFV.crearconsultaEx(strsql, cn, 1, volver)
        EY = 0
        ei = 0
        Do Until rstablas.EOF
           EY = EY + 1
           ReDim Preserve aTabla(ei)
          aTabla(ei) = rstablas(0)
           ei = ei + 1
          strsql = "Insert Into Objtabdef (codobjeto,tipobjeto,tabperfil,tabtitulo,tabsecuen) Values ("
          strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & rstablas(3) & "','"
          strsql = strsql & rstablas(4) & "'," & rstablas(5) & ")"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          rstablas.movenext
        Loop
      Call OFV.cerrarconsulta(rstablas)
        strsql = "Select * From Objtabdef Where codobjeto = " & nCodigo & " order by objid"
        Set rstablas = OFV.crearconsultaEx(strsql, cn, 1, volver)
        ei = 0
        Do Until rstablas.EOF
           ReDim Preserve nTabla(ei)
           nTabla(ei) = rstablas(0)
           ei = ei + 1
           rstablas.movenext
        Loop
      Call OFV.cerrarconsulta(rstablas)
      Call OFV.cerrarconn(cn)

     Response.write "<p><font face=arial-narrow size=1 color=#336699>  Tablas: " & EY & " </p>"

        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objtabblk Where codobjeto = " & xCodigo & " order by objid"
        Set rstablas = OFV.crearconsultaEx(strsql, cn, 1, volver)
        EY = 0
        ei = 0
        Do Until rstablas.EOF
           EY = EY + 1
           ReDim Preserve aTablab(ei)
           aTablab(ei) = rstablas(0)
           ei = ei + 1
          Ndef = 0
          For si = 0 To UBound(aTabla)
              If CStr(aTabla(si)) = CStr(rstablas(3)) Then
                 Ndef = nTabla(si)
                 Exit For
              End If
          Next

          strsql = "Insert Into Objtabblk (codobjeto,tipobjeto,tabdef,tabblksec,tabblkdescri) Values ("
          strsql = strsql & nCodigo & ",'" & sTipobjeto & "'," & Ndef & ","
          strsql = strsql & rstablas(4) & ",'" & rstablas(5) & "')"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
  'response.write strsql
          rstablas.movenext
        Loop
      Call OFV.cerrarconsulta(rstablas)
        strsql = "Select * From Objtabblk Where codobjeto = " & nCodigo & " order by objid"
        Set rstablas = OFV.crearconsultaEx(strsql, cn, 1, volver)
        ei = 0
        Do Until rstablas.EOF
           ReDim Preserve nTablab(ei)
           nTablab(ei) = rstablas(0)
           ei = ei + 1
           rstablas.movenext
        Loop
      Call OFV.cerrarconsulta(rstablas)
      Call OFV.cerrarconn(cn)

      Response.write "<p><font face=arial-narrow size=1 color=#336699>  Bloques: " & EY & " </p>"
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objtabcon Where codobjeto = " & xCodigo & " order by objid"
        Set rstablas = OFV.crearconsultaEx(strsql, cn, 1, volver)
        EY = 0
        ei = 0
        Do Until rstablas.EOF
           EY = EY + 1
           ei = ei + 1
           Nblk = 0
           For si = 0 To UBound(aTablab)
               If CStr(aTablab(si)) = CStr(rstablas(3)) Then
                  Nblk = nTablab(si)
                  Exit For
               End If
           Next

          strsql = "Insert Into Objtabcon (codobjeto,tipobjeto,tabblk,tabconsec,tabcontipo"
          strsql = strsql & ",tabconcol01,tabconcol02,tabconcol03,tabconcol04,tabconcol05"
          strsql = strsql & ",tabconcol06,tabconcol07,tabconcol08,tabconcol09,tabconcol10"
          strsql = strsql & ",tabconcol11,tabconcol12,tabconcol13,tabconcol14,tabconcol15"
          strsql = strsql & ",tabconcol16,tabconcol17,tabconcol18,tabconcol19,tabconcol20"
          strsql = strsql & ") Values ("
          strsql = strsql & nCodigo & ",'" & sTipobjeto & "'," & Nblk & ","
          strsql = strsql & rstablas(4) & ",'" & rstablas(5) & "','"
          strsql = strsql & rstablas(6) & "','" & rstablas(7) & "','" & rstablas(8) & "','"
          strsql = strsql & rstablas(9) & "','" & rstablas(10) & "','" & rstablas(11) & "','"
          strsql = strsql & rstablas(12) & "','" & rstablas(13) & "','" & rstablas(14) & "','"
          strsql = strsql & rstablas(15) & "','" & rstablas(16) & "','" & rstablas(17) & "','"
          strsql = strsql & rstablas(18) & "','" & rstablas(19) & "','" & rstablas(20) & "','"
          strsql = strsql & rstablas(21) & "','" & rstablas(22) & "','" & rstablas(23) & "','"
          strsql = strsql & rstablas(24) & "','" & rstablas(25)
          strsql = strsql & "')"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          rstablas.movenext
        Loop
      Call OFV.cerrarconsulta(rstablas)
      Call OFV.cerrarconn(cn)

      Response.write "<p><font face=arial-narrow size=1 color=#336699>  Filas: " & EY & " </p>"

      End If

  'Editor Estructurado
      If bEditor = "SI" Then
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objestr Where codobjeto = " & xCodigo & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, cn, 1, volver)
        EY = 0
        ei = 0
        Do Until rseditor.EOF
           EY = EY + 1
           ReDim Preserve aestr(ei)
           aestr(ei) = rseditor(0)
           ei = ei + 1
          strsql = "Insert Into Objestr (codobjeto,tipobjeto,contitulo,consecuen) Values ("
          strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & rseditor(3) & "',"
          strsql = strsql & rseditor(4) & ")"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
        strsql = "Select * From Objestr Where codobjeto = " & nCodigo & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, cn, 1, volver)
        ei = 0
        Do Until rseditor.EOF
           ReDim Preserve nestr(ei)
           nestr(ei) = rseditor(0)
           ei = ei + 1
           rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
     Call OFV.cerrarconn(cn)

     Response.write "<p><font face=arial-narrow size=1 color=#336699>  Estructuras: " & EY & " </p>"
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objestprty Where codobjeto = " & xCodigo & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, cn, 1, volver)
        EY = 0
        ei = 0
        Do Until rseditor.EOF
           EY = EY + 1
           ei = ei + 1
          ndetcon = 0
          For si = 0 To UBound(aestr)
              If CStr(aestr(si)) = CStr(rseditor(3)) Then
                 ndetcon = nestr(si)
                 Exit For
              End If
          Next

          strsql = "Insert Into Objestprty (codobjeto,tipobjeto,detcon,dettipo"
          strsql = strsql & ",detfont,detfsize,detfcolor,detbgcolor,detborde,detftipo"
          strsql = strsql & ",detsang,detindent1,detindent2,detvin,detvintipo,detvinindent"
          strsql = strsql & ",detnumer,detnumertipo,detnumerclase,detvinfont,detalign,dettextalign"
          strsql = strsql & ") Values ("
          strsql = strsql & nCodigo & ",'" & sTipobjeto & "'," & ndetcon & ",'" & rseditor(4) & "','"
          strsql = strsql & rseditor(5) & "','" & rseditor(6) & "','" & rseditor(7) & "','"
          strsql = strsql & rseditor(8) & "','" & rseditor(9) & "','" & rseditor(10) & "','"
          strsql = strsql & rseditor(11) & "','" & rseditor(12) & "','" & rseditor(13) & "','"
          strsql = strsql & rseditor(14) & "','" & rseditor(15) & "','" & rseditor(16) & "','"
          strsql = strsql & rseditor(17) & "','" & rseditor(18) & "','" & rseditor(19) & "','"
          strsql = strsql & rseditor(20) & "','" & rseditor(21) & "','" & rseditor(22)
          strsql = strsql & "')"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
      Call OFV.cerrarconn(cn)

      Response.write "<p><font face=arial-narrow size=1 color=#336699>  Propiedades de Estr.: " & EY & " </p>"


        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objestdet Where codobjeto = " & xCodigo
        strsql = strsql & " and (Objestdet.modi = 0 or (Objestdet.modi <> 0 "
        strsql = strsql & " and Objestdet.tipmodi <> 'E'))  order by objid"
        'Response.write "rseditor=" & strsql & "<br>"
        Set rseditor = OFV.crearconsultaEx(strsql, cn, 1, volver)
        EY = 0
        ei = 0
        Do Until rseditor.EOF
           EY = EY + 1
           ei = ei + 1
           ReDim Preserve aestdet(ei)
           aestdet(ei) = rseditor(0)
          ndetcon = 0
          For si = 0 To UBound(aestr)
              If CStr(aestr(si)) = CStr(rseditor(3)) Then
                 ndetcon = nestr(si)
                 Exit For
              End If
          Next
        ' copia de link
           If IsNumeric(rseditor("detimagen")) And rseditor("detimagen") <> "0" Then
             strsql = "SELECT * From OBJLNK where objid = '" & rseditor("detimagen") & "'"
             Set rsDsc = OFV.crearconsultaEx(strsql, cn, 1, volver)
          If Not rsDsc.EOF Then
              strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
              strsql = strsql & ") Values (" & nCodigo & ",'" & rsDsc("tipobjeto") & "',"
              strsql = strsql & "'" & rsDsc("iden1") & "','" & rsDsc("iden2") & "','" & rsDsc("iden3") & "',"
              strsql = strsql & rsDsc("codobjetoL") & ",'" & rsDsc("tipobjetoL") & "','" & rsDsc("LTipo") & "','" & rsDsc("Lurl") & "','" & rsDsc("LurlDesc") & "'"
              strsql = strsql & ")"
              'Response.write "rsdsc=" & strsql & "<br>"
              Call OFV.crearconsultaEx(strsql, cn, 1, volver)
           strsql = "SELECT * From OBJLNK where codobjeto=" & nCodigo & " and iden1='" & rsDsc("iden1") & "'"
              Set rsDsc2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
              'Response.write "rsDsc2(0)=" & strsql & " < br > "
              detimg = rsDsc2(0)
              detimgFlag = detimg
              Call OFV.cerrarconsulta(rsDsc2)
             End If
            Call OFV.cerrarconsulta(rsDsc)
            Else
             detimg = rseditor(9)
                detimgFlag = ""
           End If

            '
       strsql = "Insert Into Objestdet (codobjeto,tipobjeto,detcon,detsecuen,dettipo"
          strsql = strsql & ",detdepen,dettitulo,dettexto,detimagen,dettabla"
          strsql = strsql & ",detfont,detfsize,detfcolor,detbgcolor,detborde,detftipo"
          strsql = strsql & ",detsang,detindent1,detindent2,detvin,detvintipo,detvinindent"
          strsql = strsql & ",detnumer,detnumertipo,detnumerclase,detvinfont,detalign,dettextalign"
          strsql = strsql & ",detdefault,modi"
          strsql = strsql & ") Values ("
          strsql = strsql & nCodigo & ",'" & sTipobjeto & "'," & ndetcon & "," & rseditor(4)
          strsql = strsql & ",'" & rseditor(5) & "',"
          strsql = strsql & rseditor(6) & ",'" & rseditor(7) & "','" & rseditor(8) & "','"
          strsql = strsql & detimg & "','" & rseditor(10) & "','"
          strsql = strsql & rseditor(11) & "','" & rseditor(12) & "','" & rseditor(13) & "','"
          strsql = strsql & rseditor(14) & "','" & rseditor(15) & "','" & rseditor(16) & "','"
          strsql = strsql & rseditor(17) & "','" & rseditor(18) & "','" & rseditor(19) & "','"
          strsql = strsql & rseditor(20) & "','" & rseditor(21) & "','" & rseditor(22) & "','"
          strsql = strsql & rseditor(23) & "','" & rseditor(24) & "','" & rseditor(25) & "','"
          strsql = strsql & rseditor(26) & "','" & rseditor(27) & "','" & rseditor(28) & "','"
          strsql = strsql & rseditor(31)
          strsql = strsql & "',0)"
          'Response.write "rseditor=" & strsql & "<br>"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)


          ' cambio el iden1 y iden2
            If detimgFlag <> "" Then
             strsql = "select objid from objestdet where detimagen='" & detimgFlag & "' and codobjeto=" & nCodigo
              '  Response.write "ttttt=" & strsql & "<br>"
              Set rsDsc2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
              strsql = "update objlnk set iden1=" & rsDsc2(0) & ",iden2=" & rsDsc2(0) & " where objid=" & detimgFlag & " and codobjeto=" & nCodigo
              'Response.write "detimgFlag=" & strsql & "<br>"
              Set rsDsc = OFV.crearconsultaEx(strsql, cn, 1, volver)
              Call OFV.cerrarconsulta(rsDsc)
              Call OFV.cerrarconsulta(rsDsc2)
            End If

            rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
      Call OFV.cerrarconn(cn)

        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objestdet Where codobjeto = " & nCodigo & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, cn, 1, volver)
       ' Response.write "reditor2=" & strsql & "<br>"
        EY = 0
        ei = 0
        Do Until rseditor.EOF
           EY = EY + 1
           ei = ei + 1
         '  Response.write "rseditor entro <br>"
           ReDim Preserve nestdet(ei)
           nestdet(ei) = rseditor(0)
          ndetdep = 0
          If rseditor(5) <> "A" Then
             For si = 0 To UBound(nestdet)
                 If CStr(aestdet(si)) = CStr(rseditor(6)) Then
                    ndetdep = nestdet(si)
                    Exit For
                 End If
             Next
          End If

          strsql = "update Objestdet set detdepen = '" & ndetdep & "' where objid = " & rseditor(0)
        '  Response.write "rseditor3=" & strsql & "<br>"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
      Call OFV.cerrarconn(cn)
      Response.write "<p><font face=arial-narrow size=1 color=#336699>  Contenido de Estr.: " & EY & " </p>"

      End If


  'Procedimientos
      If bProcedimiento = "SI" Then
         ReDim Preserve ProcOld(0)
         ReDim Preserve ProcNew(0)
         ReDim Preserve AccPnew(0)
         ReDim Preserve AccIOld(0)
         ReDim Preserve AccINew(0)


        Response.write "<p><font face=arial-narrow size=1 color=#336699>  Procedimientos </p>"
        ntime1 = Minute(Now) & ":" & Second(Now)
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objproced Where codobjeto = " & xCodigo & " order by codproced "
        Set rsProced = OFV.crearconsultaEx(strsql, cn, 1, volver)
        i = 0
        inp = 0
        inpx = 0
        ip = 1
        Do Until rsProced.EOF
          inpx = inpx + 1
          If inpx > 50 Then
             inpx = 0
             Response.write "tiempo  parcial: " & Minute(Now) & ":" & Second(Now) & "<br>"
          End If

          ReDim Preserve ProcOld(i)
          ReDim Preserve ProcNew(i)
          inp = inp + 1
          nNewProced = rsProced(7)
          nProced = GTkey(nCodigo & inp)
          ProcOld(i) = rsProced(7)
          ProcNew(i) = nProced
          Neven = 0
          For ei = 0 To UBound(aevento)
              If CStr(aevento(ei)) = CStr(rsProced(3)) Then
                 Neven = nevento(ei)
                 Exit For
              End If
          Next
          Nsect = 0
          For si = 0 To UBound(asector)
              If CStr(asector(si)) = CStr(rsProced(4)) Then
                 Nsect = nsector(si)
                 Exit For
              End If
          Next

          strsql = "Insert Into Objproced (codobjeto,tipobjeto,codeven,codsector,codperfil,"
          strsql = strsql & "secuencia,codproced,frecuencia,referencia) Values (" & nCodigo & ",'" & sTipobjeto & "',"
          strsql = strsql & Neven & "," & Nsect & ",'" & rsProced(5) & "',"
          strsql = strsql & rsProced(6) & ",'" & nProced & "','" & rsProced(8) & "','" & rsProced(9) & "')"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          rsProced.movenext
          i = i + 1
          ip = ip + 1
        Loop
        Call OFV.cerrarconsulta(rsProced)
        Call OFV.cerrarconn(cn)
        Response.write "<p><font face=arial-narrow size=1 color=#336699> Procedimientos: " & inp & " </p>"
  '-----------------------------------------------
  '  Procedimientos - Acciones
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objaccion Where codobjeto = " & xCodigo
        strsql = strsql & " and (OBJACCION.modi = 0 or (OBJACCION.modi <> 0 "
        strsql = strsql & " and OBJACCION.tipmodi <> 'E'))  order by codproced , codaccion"
        Set rs1 = OFV.crearconsultaEx(strsql, cn, 1, volver)

        ind = 0
        acccn = 0
        iaa = 0
        iaap = 0
        Do Until rs1.EOF
          inpx = inpx + 1
          If inpx > 50 Then
             inpx = 0
             Response.write "tiempo  parcial: " & Minute(Now) & ":" & Second(Now) & "<br>"
          End If

           ReDim Preserve AccPnew(iaa)
           ReDim Preserve AccIOld(iaa)
           ReDim Preserve AccINew(iaa)

           ind = ind + 1
           sukey = GTkey(nCodigo & ind)

           For iaap = iaap To UBound(ProcOld)
               If CStr(ProcOld(iaap)) = CStr(rs1(3)) Then
                  nProced = ProcNew(iaap)
                  Exit For
               End If
               If CStr(ProcOld(iaap)) > CStr(rs1(3)) Then
                  nProced = "XXXX"
                  Exit For
               End If

           Next
           If nProced <> "XXXX" And iaap <= UBound(ProcOld) Then
           AccIOld(iaa) = rs1(18)
           AccINew(iaa) = sukey
           AccPnew(iaa) = ProcNew(iaap)

'''''''''''

       If rs1(19) > 0 Then
          strsql = "SELECT * From OBJLNK where objid = " & rs1(19)
          Set rsDsc = OFV.crearconsultaEx(strsql, cn, 1, volver)
          If Not rsDsc.EOF Then


                strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL "
                strsql = strsql & ") Values (" & nCodigo & ",'" & sTipobjeto & "',"
                strsql = strsql & "'" & nProced & "','" & sukey & "',"
                strsql = strsql & rsDsc(6) & ",'" & rsDsc(7) & "'"
                strsql = strsql & ")"
                Call OFV.crearconsultaEx(strsql, cn, 1, volver)

           Response.write "<p><font face=arial-narrow size=1 color=#336699>  Procedimiento: acciones links </p>"


          End If
          Call OFV.cerrarconsulta(rsDsc)

        strsql = "SELECT objid From OBJLNK where iden2 = '" & sukey & "' "
        Set rsDsc = OFV.crearconsultaEx(strsql, cn, 1, volver)
        If Not rsDsc.EOF Then
           wsval3 = rsDsc(0)
        Else
           wsval3 = 0
        End If
        Call OFV.cerrarconsulta(rsDsc)

        strsql = "Insert into Objaccion (Codobjeto,Tipobjeto,Codproced,Tipoaccion,modo,codcomp,"
        strsql = strsql & "tipocomp,sectorcont,perfilcont,secuen,referencia,tiempo,copias,"
        strsql = strsql & "referencia2,referencia3,codaccion,rellink,reflink) Values("
        strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & nProced & "'," & rs1(4) & ",'" & rs1(5) & "',"
        strsql = strsql & rs1(6) & ",'" & rs1(7) & "'," & rs1(8) & ",'" & rs1(9) & "'," & rs1(10) & ",'"
        strsql = strsql & rs1(11) & "'," & rs1(12) & "," & rs1(13)
        strsql = strsql & ",'" & rs1(16) & "','" & rs1(17) & "','" & sukey & "'," & wsval3 & ",'" & rs1(20) & "')"

        Call OFV.crearconsultaEx(strsql, cn, 1, volver)

    Else

        strsql = "Insert into Objaccion (Codobjeto,Tipobjeto,Codproced,Tipoaccion,modo,codcomp,"
        strsql = strsql & "tipocomp,sectorcont,perfilcont,secuen,referencia,tiempo,copias,"
        strsql = strsql & "referencia2,referencia3,codaccion,rellink,reflink) Values("
        strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & nProced & "'," & rs1(4) & ",'" & rs1(5) & "',"
        strsql = strsql & rs1(6) & ",'" & rs1(7) & "'," & rs1(8) & ",'" & rs1(9) & "'," & rs1(10) & ",'"
        strsql = strsql & rs1(11) & "'," & rs1(12) & "," & rs1(13)
        strsql = strsql & ",'" & rs1(16) & "','" & rs1(17) & "','" & sukey & "',0,'')"

        Call OFV.crearconsultaEx(strsql, cn, 1, volver)

    End If



'''''''''''''''





           iaa = iaa + 1
           End If
           rs1.movenext
        Loop
        Call OFV.cerrarconsulta(rs1)
        Call OFV.cerrarconn(cn)
      Response.write "<p><font face=arial-narrow size=1 color=#336699>  Procedimiento: acciones " & ind & "</p>"

  '-----------------------------------------------
  '  Procedimientos - Acciones Alternativas
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "SELECT OBJACCIONCONTROL.* FROM OBJACCION, OBJACCIONCONTROL where "
        strsql = strsql & " OBJACCION.codaccion = OBJACCIONCONTROL.idaccion "
        strsql = strsql & " and OBJACCION.codobjeto = " & xCodigo
        strsql = strsql & " and (OBJACCIONcontrol.modi = 0 or (OBJACCIONcontrol.modi <> 0"
        strsql = strsql & " and OBJACCIONcontrol.tipmodi <> 'E')) order by OBJACCION.codaccion"

        Set rsAlt = OFV.crearconsultaEx(strsql, cn, 1, volver)

        Do Until rsAlt.EOF
          inpx = inpx + 1
          If inpx > 50 Then
             inpx = 0
             Response.write "tiempo  parcial: " & Minute(Now) & ":" & Second(Now) & "<br>"
          End If

           For iaap = 0 To UBound(AccIOld)
               If CStr(AccIOld(iaap)) = CStr(rsAlt(3)) Then
                  nProced = AccPnew(iaap)
                  sukey = AccINew(iaap)
                  Exit For
               End If
           Next


           nconti = "0"
           If iaap <= UBound(AccIOld) Then

              If Len(rsAlt(6)) > 0 Then
                  For iaa = 0 To UBound(AccIOld)
                      If CStr(AccIOld(iaa)) = CStr(rsAlt(6)) Then
                         nconti = AccINew(iaa)
                         Exit For
                      End If
                  Next
              End If

           acccn = acccn + 1

           strsql = "Insert Into Objaccioncontrol (codproced,accion,idaccion,secuencia,referencia,"
           strsql = strsql & "continuador) Values ('" & nProced & "'," & rsAlt(2) & ",'" & sukey
           strsql = strsql & "'," & rsAlt(4) & ",'" & rsAlt(5) & "','" & nconti & "')"
           Call OFV.crearconsultaEx(strsql, cn, 1, volver)
           End If
           rsAlt.movenext

        Loop
        Call OFV.cerrarconsulta(rsAlt)

      Response.write "<p><font face=arial-narrow size=1 color=#336699>  Procedimiento: acciones alt " & acccn & "</p>"

  '-----------------------------------------------

        Call OFV.cerrarconn(cn)

        Response.write "tiempo  fin: " & Minute(Now) & ":" & Second(Now) & "<br>"

      End If
  'Controles
      If bControl = "SI" Then
         Dim Ocontrol()
         Dim ycontrol()
         CY = 0
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objcontrol Where codobjeto = " & xCodigo & " order by objid"
        Set rsControl = OFV.crearconsultaEx(strsql, cn, 1, volver)
        Do Until rsControl.EOF
          strsql = "Insert Into Objcontrol (codobjeto,tipobjeto,codsector,frecuencia,codperfil,referencia) Values ("
          strsql = strsql & nCodigo & ",'" & sTipobjeto & "'," & rsControl(3) & ",'"
          strsql = strsql & rsControl(4) & "','" & rsControl(5) & "','" & rsControl(6) & "')"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          ReDim Preserve Ocontrol(CY)
          Ocontrol(CY) = rsControl(0)
          CY = CY + 1
          rsControl.movenext
        Loop
        Call OFV.cerrarconsulta(rsControl)
        CY = 0
        strsql = "Select objid From Objcontrol Where codobjeto = " & nCodigo & " order by objid"
        Set rsControlb = OFV.crearconsultaEx(strsql, cn, 1, volver)
        Do Until rsControlb.EOF
          ReDim Preserve ycontrol(CY)
          ycontrol(CY) = rsControlb(0)
          CY = CY + 1
          rsControlb.movenext
        Loop
        Call OFV.cerrarconsulta(rsControlb)

          strsql = "Select * From Objcontrol2 Where codobjeto = " & xCodigo
          strsql = strsql & " and (Objcontrol2.modi = 0 or (Objcontrol2.modi <> 0"
          strsql = strsql & " and Objcontrol2.tipmodi <> 'E')) order by control"
          Set rsControl2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
          Do Until rsControl2.EOF
             nControl = 0
             For CY = 0 To UBound(ycontrol)
                 If CStr(rsControl2("control")) = CStr(Ocontrol(CY)) Then
                    nControl = ycontrol(CY)
                 End If
             Next
            strsql = "Insert Into Objcontrol2 (codobjeto,tipobjeto,control,descri,secuencia) Values ("
            strsql = strsql & nCodigo & ",'" & sTipobjeto & "'," & nControl & ",'" & rsControl2(4)
            strsql = strsql & "'," & rsControl2(5) & ")"
            Call OFV.crearconsultaEx(strsql, cn, 1, volver)
            rsControl2.movenext
          Loop
          Call OFV.cerrarconsulta(rsControl2)


      Call OFV.cerrarconn(cn)
      End If

  'Esquemas
      If bEsquema = "SI" Then
         Dim Oesque()
         Dim Nesque()
         CZ = 0
    Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objesquema Where codobjeto = " & xCodigo & " order by objid"
        Set rsEsquema = OFV.crearconsultaEx(strsql, cn, 1, volver)
        i = 0
        Do Until rsEsquema.EOF
          Neven = 0
          For ei = 0 To UBound(aevento)
              If CStr(aevento(ei)) = CStr(rsEsquema(3)) Then
                 Neven = nevento(ei)
                 Exit For
              End If
          Next
          strsql = "Insert Into Objesquema (codobjeto,tipobjeto,codevento,sectororigen,"
          strsql = strsql & "sectordestino,refesq) Values (" & nCodigo & ",'" & sTipobjeto & "',"
          strsql = strsql & Neven & "," & rsEsquema(4) & "," & rsEsquema(5) & ",'" & rsEsquema(6) & "')"
          Call OFV.crearconsultaEx(strsql, cn, 1, volver)
          ReDim Preserve Oesque(CZ)
          Oesque(CZ) = rsEsquema(0)
          CZ = CZ + 1
          rsEsquema.movenext
          i = i + 1
          Loop
          Call OFV.cerrarconsulta(rsEsquema)
          CZ = 0
          strsql = "Select objid From Objesquema Where codobjeto = " & nCodigo & " order by objid"
          Set rsEsquemab = OFV.crearconsultaEx(strsql, cn, 1, volver)
          Do Until rsEsquemab.EOF
          ReDim Preserve Nesque(CZ)
          Nesque(CZ) = rsEsquemab(0)
          CZ = CZ + 1
          rsEsquemab.movenext
          Loop
          Call OFV.cerrarconsulta(rsEsquemab)
          strsql = "Select * From Objesquema2 Where codobjeto = " & xCodigo
          strsql = strsql & " and (Objesquema2.modi = 0 or (Objesquema2.modi <> 0"
          strsql = strsql & " and Objesquema2.tipmodi <> 'E')) order by codesquema "
          Set rsEsquema2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
          Do Until rsEsquema2.EOF
             nEsquema = 0
             For CZ = 0 To UBound(Nesque)
                 If CStr(rsEsquema2("codesquema")) = CStr(Oesque(CZ)) Then
                    nEsquema = Nesque(CZ)
                 End If
             Next
            strsql = "Insert Into Objesquema2 (codobjeto,tipobjeto,codesquema,codcuenta,sentido,refesq2) "
            strsql = strsql & "Values (" & nCodigo & ",'" & sTipobjeto & "'," & nEsquema & ","
            strsql = strsql & rsEsquema2(4) & ",'" & rsEsquema2(5) & "','" & rsEsquema2(6) & "')"
            Call OFV.crearconsultaEx(strsql, cn, 1, volver)
            rsEsquema2.movenext
          Loop
          Call OFV.cerrarconsulta(rsEsquema2)


      Call OFV.cerrarconn(cn)
      End If

 'Consideraciones Generales
  ''''''''''''''
    If bMain = "SI" Then
        Set cn = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objmain Where codobjeto = " & xCodigo
       ' Response.write "bMain=" & strsql & "<br>"
        Set rsObjetivos = OFV.crearconsultaEx(strsql, cn, 1, volver)
        Do Until rsObjetivos.EOF
            strsql = "Insert Into Objmain (codobjeto,tipobjeto,archivo,titulo) Values ("
            strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & rsObjetivos(3) & "','"
            strsql = strsql & rsObjetivos(4) & "')"
           ' Response.write "rsObjetivos=" & strsql & "<br>"
         Call OFV.crearconsultaEx(strsql, cn, 1, volver)
           ' rsEvento.movenext
            rsObjetivos.movenext
        Loop
        Call OFV.cerrarconsulta(rsObjetivos)

        strsql = "Select * From Objalcance Where codobjeto = " & xCodigo

        Set rsObjetivos = OFV.crearconsultaEx(strsql, cn, 1, volver)
        Do Until rsObjetivos.EOF
            strsql = "Insert Into Objalcance (codobjeto,tipobjeto,archivo,titulo) Values ("
            strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & rsObjetivos(3) & "','"
            strsql = strsql & rsObjetivos(4) & "')"
            Call OFV.crearconsultaEx(strsql, cn, 1, volver)
            'rsEvento.movenext
            rsObjetivos.movenext
        Loop
        Call OFV.cerrarconsulta(rsObjetivos)

        strsql = "Select * From Objconsid Where codobjeto = " & xCodigo
        Set rsObjetivos = OFV.crearconsultaEx(strsql, cn, 1, volver)
        Do Until rsObjetivos.EOF
            strsql = "Insert Into Objconsid (codobjeto,tipobjeto,archivo,titulo) Values ("
            strsql = strsql & nCodigo & ",'" & sTipobjeto & "','" & rsObjetivos(3) & "','"
            strsql = strsql & rsObjetivos(4) & "')"
            Call OFV.crearconsultaEx(strsql, cn, 1, volver)
            'rsEvento.movenext
            rsObjetivos.movenext
        Loop
        Call OFV.cerrarconsulta(rsObjetivos)


      On Error Resume Next
      Set fso = Server.CreateObject("scripting.filesystemobject")
      sPath = Server.mappath(Request.servervariables("SCRIPT_NAME"))
      sPath = Mid(sPath, 1, InStrRev(sPath, "\", -1))
      sPathObj = sPath & "CarpetaObjetivos\"
      sPathAlc = sPath & "CarpetaAlcance\"
      sPathCon = sPath & "CarpetaConsideracion\"

'      Set cn = OFV.conectar(OFV.strconn2)

      strsql = "Select archivo From Objmain Where codobjeto = " & xCodigo
      Set rsObjetivos = OFV.crearconsultaEx(strsql, cn, 1, volver)
      If Not rsObjetivos.EOF Then
        sOrigen = sPathObj & rsObjetivos(0)
        sDestino = sPathObj & "Obj" & CStr(nCodigo) & ".rtf"
        fso.CopyFile sOrigen, sDestino
      End If
      Call OFV.cerrarconsulta(rsObjetivos)

      strsql = "Select archivo From Objalcance Where codobjeto = " & xCodigo
      Set rsAlcances = OFV.crearconsultaEx(strsql, cn, 1, volver)
      If Not rsAlcances.EOF Then
        sOrigen = sPathAlc & rsAlcances(0)
        sDestino = sPathAlc & "Alc" & CStr(nCodigo) & ".rtf"
        fso.CopyFile sOrigen, sDestino
      End If
      Call OFV.cerrarconsulta(rsAlcances)

      strsql = "Select archivo From Objconsid Where codobjeto = " & xCodigo
      Set rsConsid = OFV.crearconsultaEx(strsql, cn, 1, volver)
      If Not rsConsid.EOF Then
        sOrigen = sPathCon & rsConsid(0)
        sDestino = sPathCon & "Con" & CStr(nCodigo) & ".rtf"
        fso.CopyFile sOrigen, sDestino
      End If
      On Error Resume Next
      Call OFV.cerrarconsulta(rsConsid)
      Call OFV.cerrarconn(cn)
      Response.write "tiempo inicial: " & ntime0 & " final: " & Minute(Now) & ":" & Second(Now)

    End If


  Response.write "<center>"

      sInput = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "")
      Response.write OFV.GenForm("Copiar", "", "vuelve.asp", "", "post", "", sInput, "si")
  Response.write "</center>"
      sHTML = "</body></html>"


    Else
      sHTML = OFV.FormBody
   'Objetos
    Set cn = OFV.conectar(OFV.strconn2)
       strsql = "select * from estados"
       Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
       inp = 0
       Do Until rsTipo.EOF
          ReDim Preserve aestr(inp)
          ReDim Preserve nestr(inp)
          aestr(inp) = UCase(rsTipo(0))
          nestr(inp) = UCase(rsTipo(1))
          inp = inp + 1
       rsTipo.movenext
       Loop

       strsql = "select tipobjeto,descri from subtema"
       Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
       inp = 0
       Do Until rsTipo.EOF
          ReDim Preserve asector(inp)
          ReDim Preserve nsector(inp)
          asector(inp) = UCase(rsTipo(0))
          nsector(inp) = UCase(rsTipo(1))
          inp = inp + 1
       rsTipo.movenext
       Loop

       strsql = "select tipobjeto,descri from tipobjeto"
       Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
       inp = 0
       Do Until rsTipo.EOF
          ReDim Preserve aevento(inp)
          ReDim Preserve nevento(inp)
          aevento(inp) = UCase(rsTipo(0))
          nevento(inp) = UCase(rsTipo(1))
          inp = inp + 1
       rsTipo.movenext
       Loop
       Call OFV.cerrarconsulta(rsTipo)

       strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
       strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and objetos.codobjeto = " & nCodigo

       Set rsObj = OFV.crearconsultaEx(strsql, cn, 1, volver)

       strsql = "Select codobjmaster,tipobjmaster From relaobj Where codojasoc = " & nCodigo
       strsql = strsql & " and tipobjasoc = '" & rsObj(2) & "' "
       Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)



       bSector = "S/D"
       bControl = "N"
       Nsect = 0
       If Not rsNew.EOF Then
       Nsect = rsNew(0)
       For inp = 0 To UBound(asector)
           If asector(inp) = UCase(rsNew(1)) Then
              bSector = nsector(inp)
              bControl = "S"
              Exit For
            End If
       Next

       If bControl = "N" Then
          For inp = 0 To UBound(aevento)
           If aevento(inp) = UCase(rsNew(1)) Then
              bSector = nevento(inp)
              bControl = "O"
              Exit For
            End If
          Next
       End If

       If bControl = "O" Then
          strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
          strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and objetos.codobjeto = " & rsNew(0)

          Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
          If Not rsNew.EOF Then
             bSector = bSector & ": " & rsNew(1) & " - " & rsNew(3)
             For inp = 0 To UBound(aestr)
                 If aestr(inp) = UCase(rsNew(9)) Then
                    bSector = bSector & " (" & nestr(inp) & ")"
                    Exit For
                 End If
             Next

          End If
       Else
          strsql = "Select Prodserv.*,Tema.descri From Tema, Prodserv where "
          strsql = strsql & "Tema.codtema = Prodserv.codtema and prodserv.codprodserv = " & rsNew(0)
          Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
          If Not rsNew.EOF Then
             bSector = bSector & ": " & rsNew(5) & "/" & rsNew(1)
          End If

       End If

       Call OFV.cerrarconsulta(rsNew)
       End If

       bEvento = "S/D"
       For inp = 0 To UBound(aevento)
           If aevento(inp) = UCase(rsObj(2)) Then
              bEvento = nevento(inp)
              Exit For
            End If
       Next

       'origen
        Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Copiar:", "si")
        sInput = bEvento & ": " & rsObj(1) & " - " & rsObj(3)
        For inp = 0 To UBound(aestr)
            If aestr(inp) = UCase(rsObj(9)) Then
               sInput = sInput & " (" & nestr(inp) & ")"
               Exit For
            End If
        Next
        Celda = Celda & OFV.GenCelda("", "cc", "", "", "", "", "", sInput, "si")
        Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

       'Superior
        Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Perteneciente a:", "si")
        sInput = bSector
        Celda = Celda & OFV.GenCelda("", "cc", "", "", "", "", "", sInput, "si")
        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

        'Como
        Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Como:", "si")
        sInput = sOpcion
        For inp = 0 To UBound(aevento)
            If aevento(inp) = UCase(sOpcion) Then
               sInput = nevento(inp)
               Exit For
            End If
        Next
        Celda = Celda & OFV.GenCelda("", "cc", "", "", "", "", "", sInput, "si")
        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'Descripcion
      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Descripcion", "si")
      sInput = OFV.GenerarInput(rsObj(3).Name, rsObj(3), "agregar", "Copiar", "80", "dt", "")
      Celda = Celda & OFV.GenCelda("", "cc", "", "", "", "", "", sInput, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'Identificacion
      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Identificacion", "si")
      sInput = OFV.GenerarInput(rsObj(1).Name, rsObj(1), "agregar", "copiar", "20", "dt", "")
      Celda = Celda & OFV.GenCelda("", "cc", "", "", "", "", "", sInput, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      strsql = "Select nivel From Tipobjeto Where tipobjeto = '" & sOpcion & "'"
      Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
      sTipo = rsTipo(0)
      Call OFV.cerrarconsulta(rsTipo)
      Select Case sTipo
      Case "C"
          strsql = "Select tipobjeto From Subtema Where levc = -1"
      Case "D"
          strsql = "Select tipobjeto From Subtema Where levd = -1"
      Case "E"
          strsql = "Select tipobjeto From Subtema Where leve = -1"
      Case "F"
          strsql = "Select tipobjeto From Subtema Where levf = -1"
      Case "G"
          strsql = "Select tipobjeto From Subtema Where levg = -1"
      Case "H"
          strsql = "Select tipobjeto From Subtema Where levh = -1"
      Case "I"
          strsql = "Select tipobjeto From Subtema Where levi = -1"
      Case "J"
          strsql = "Select tipobjeto From Subtema Where levj = -1"
      Case "K"
          strsql = "Select tipobjeto From Subtema Where levk = -1"
      Case "L"
          strsql = "Select tipobjeto From Subtema Where levl = -1"
      End Select
      Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
      aTipos0 = "('"
      Do Until rsTipo0.EOF
        aTipos0 = aTipos0 & rsTipo0(0) & "','"
        rsTipo0.movenext
      Loop
      aTipos0 = Left(aTipos0, Len(aTipos0) - 2) & ")"

      Select Case sTipo
      Case "C"
          strsql = "Select tipobjeto From Tipobjeto Where levc = -1"
      Case "D"
          strsql = "Select tipobjeto From Tipobjeto Where levd = -1"
      Case "E"
          strsql = "Select tipobjeto From Tipobjeto Where leve = -1"
      Case "F"
          strsql = "Select tipobjeto From Tipobjeto Where levf = -1"
      Case "G"
          strsql = "Select tipobjeto From Tipobjeto Where levg = -1"
      Case "H"
          strsql = "Select tipobjeto From Tipobjeto Where levh = -1"
      Case "I"
          strsql = "Select tipobjeto From Tipobjeto Where levi = -1"
      Case "J"
          strsql = "Select tipobjeto From Tipobjeto Where levj = -1"
      Case "K"
          strsql = "Select tipobjeto From Tipobjeto Where levk = -1"
      Case "L"
          strsql = "Select tipobjeto From Tipobjeto Where levl = -1"
      End Select
      Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
      aTipos = "('"
      Do Until rsTipo.EOF
        aTipos = aTipos & rsTipo(0) & "','"
        rsTipo.movenext
      Loop
      aTipos = Left(aTipos, Len(aTipos) - 2) & ")"
      If aTipos <> ")" Then
        strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
        strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and (Objprty.estado = 1 "
        strsql = strsql & "Or Objprty.estado = 3) And Objetos.tipobjeto In " & aTipos
        Set rsTipo2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
      End If
      If aTipos0 <> ")" Then
        strsql = "Select Prodserv.*,Tema.descri From Tema, Prodserv where "
        strsql = strsql & "Tema.codtema = Prodserv.codtema and Prodserv.marcaps In " & aTipos0
        Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
      End If

      'Esquema superior
      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Esquema Superior", "si")
      sInput = "<font class=at><select name=Esquema>"
      If aTipos <> ")" Then
        Do Until rsTipo2.EOF
          sInput = sInput & "<option value=" & rsTipo2(0) & "_" & rsTipo2(2)
             If bControl = "O" Then
                If CStr(Nsect) = CStr(rsTipo2(0)) Then
                   sInput = sInput & " selected "
                End If
             End If
             sInput = sInput & " >"
             For inp = 0 To UBound(aevento)
                 If aevento(inp) = UCase(rsTipo2(2)) Then
                    sInput = sInput & nevento(inp) & ": "
                    Exit For
                 End If
             Next
             sInput = sInput & rsTipo2(1)
         sInput = sInput & " - " & rsTipo2(3)
             For inp = 0 To UBound(aestr)
                 If aestr(inp) = UCase(rsTipo2(9)) Then
                    sInput = sInput & " (" & nestr(inp) & ")"
                    Exit For
                 End If
             Next

          rsTipo2.movenext
        Loop
      End If
      If aTipos0 <> ")" Then
        Do Until rsTipo0.EOF
          sInput = sInput & "<option value=" & rsTipo0(0) & "_" & rsTipo0(4)
             If bControl = "S" Then
                If CStr(Nsect) = CStr(rsTipo0(0)) Then
                   sInput = sInput & " selected "
                End If
             End If
             sInput = sInput & " >"
             For inp = 0 To UBound(asector)
                 If asector(inp) = UCase(rsTipo0(4)) Then
                    sInput = sInput & nsector(inp) & ": "
                    Exit For
                 End If
             Next
             sInput = sInput & rsTipo0(5)
          sInput = sInput & "/" & rsTipo0(1) & "/"
          rsTipo0.movenext
        Loop
      End If
      sInput = sInput & "</select></font>"
      Celda = Celda & OFV.GenCelda("", "cc", "", "", "", "", "", sInput, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      Tabla = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")

      If aTipos <> ")" Then Call OFV.cerrarconsulta(rsTipo2)
      If aTipos0 <> ")" Then Call OFV.cerrarconsulta(rsTipo0)

      Tabla = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")

      Fila = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

      sInput = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "")
      Celda = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
      sAccion = "Copiar2.asp?Codigo=" & nCodigo & "&Tipo=" & sTipo & "&Perfil=" & sPerfil
      sAccion = sAccion & "&Opcion=" & sOpcion & "&Copiar=SI"
      Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

      sFunc = "window.close()"
      sInput = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
      Celda = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
      Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
      Fila = Formulario & OFV.GenFRow()

      Tabla = Fila & OFV.GenFTabla()

      sHTML = sHTML & "<br><br><br><br><center>" & Tabla & "</center></body></html>"
      Call OFV.cerrarconsulta(rsObj)
      Call OFV.cerrarconn(cn)
    End If



'End If


     Generar = sHTML


End Function

 Function GTkey(XCodobJ)

  Dim GThms, GTses, GTsecu, GTsecuf, GTi, Gtcs

  GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
  GTses = Session.sessionid
  GTsecu = CStr(XCodobJ & GThms & GTses)

  GTsecuf = ""

  For GTi = 1 To Len(GTsecu)
  Gtcs = Mid(GTsecu, GTi, 1)
  If Gtcs >= "0" And Gtcs <= "9" Then
     GTsecuf = GTsecuf & Gtcs
End If
  Next

  GTkey = Trim(GTsecuf)

End Function

