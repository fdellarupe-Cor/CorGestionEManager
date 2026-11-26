<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorCopyGen.inc" -->
<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  XXXX = NEWEMPRE("NO")

  nCodigo    = Request.querystring("Codigo")
  sTipo      = Request.querystring("Tipo")

  nCodESQS    = Request.querystring("CodESQS")
  sTipESQS    = Request.querystring("TipESQS")

  nCodESQSO    = Request.querystring("CodESQSO")
  sTipESQSO    = Request.querystring("TipESQSO")

  bCopiar    = Request.querystring("Copiar")
  sOpcion    = Request.querystring("Opcion")

  SPERIODO    = Request.querystring("PERIODO")

  sUsuario   = Session("Usuario")
  sTipobjeto = sTipo
  volver = Request.querystring("volver")

  VOLVER = REPLACE((VOLVER & ""),"../","")
  VOLVER = REPLACE((VOLVER & ""),"_","&")

'''''''''''''

  CCGIOCOD = nCodigo
  CCGIOTIP = sTipo

  CCGINCOD = 0
  CCGINTIP = sTipobjeto

  CCGICESQ = nCodESQS
  CCGITESQ = sTipESQS

  CCGOCESQ = nCodESQSO
  CCGOTESQ = sTipESQSO



        Response.write OFV.FormBody

        Response.write "<font face='verdana' size=3 color=#336699>"




 if sopcion = "TNPPL" OR sopcion = "TNP" OR SOPCION = "TNPDA" OR SOPCION = "TNPDAV" THEN



        Response.write "<center><p style='border-style=solid;border-width:1;border-color:#888888;COLOR:#888888;BACKGROUND-COLOR:#DDDDDD;' >"
        Response.write "<b>  ACTIVACION DE CIRCUITO </b></p><br>"

        styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; "
        styl = styl & "SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa;"
        styl = styl & "SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"



        Response.write "<font face='verdana' size=2 color=#336699>"

        ntime0     = Minute(Now) & ":" & Second(Now)
        Response.write "tiempo proc inicio: " & ntime0


 ELSE



        Response.write "<center><p style='border-style=solid;border-width:1;border-color:#eeeeee;' ><b>  REEMPLAZO DEL PROCESO/SUBPROCESO </b></p><br>"

        styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; "
        styl = styl & "SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa;"
        styl = styl & "SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"



        Response.write "<font face='verdana' size=2 color=#336699>"

        ntime0     = Minute(Now) & ":" & Second(Now)
        Response.write "tiempo proc inicio: " & ntime0



    Set cn = OFV.conectar(OFV.strconn2)




  Strsql = "Select * From OBJETOS  where CODOBJETO = " & CCGIOCOD
  Strsql = Strsql & " and tipobjeto = '" & CCGIOTIP & "'  "

 '' response.write strsql & "<br>"
  set RSOBJIN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT RSOBJIN.eof then


     CCGIEMPR = RSOBJIN(7)
     CCGINPRF = "OBJMAIN"
     CCGIDESC = RSOBJIN(3)
     CCGICDID = RSOBJIN(1)

     CCGICREL = RSOBJIN(15)

  END IF


'' sopcion = "TNP"                      ' GENERA ENTRADA PARA EL NUEVO PERIODO - DESDE CERO
'' SOPCION = "TNPDA"                    ' GENERA NUEVO PERIODO - TOMA CIRCUITO DESDE PERIODO ANTERIOR
'' SOPCION = "TNPDAV"                   ' GENERA NUEVO PERIODO - TOMA CIRCUITO DESDE PERIODO ANTERIOR DE LA VERSION ANTERIOR DEL DOC.VINCULADO







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' VERIFICA EL PERFIL DEL DOCUMENTO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          bSector        = "NO"
          bEvento        = "NO"
          bProcedimiento = "NO"
          bControl       = "NO"
          bEsquema       = "NO"
          bTablas        = "NO"
          bDocum         = "NO"
          bEditor        = "NO"
          bDescri        = "NO"
          bMain          = "NO"

       strsql = "Select * From Objperfildoc Where codobjeto =" & CCGIOCOD
'  response.write "sql " & strsql
       Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)
       If rsPerfil.EOF Then
          perfilsp = "N"
'  response.write "no "

          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '''' VERIFICA EL PERFIL GENERAL DEL OBJETO ''''''
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


          strsql = "Select Objperfil.* From Objperfil, Tipobjeto where "
          strsql = strsql & "Objperfil.tipoperfil = Tipobjeto.perfil "
          strsql = strsql & "and Tipobjeto.tipobjeto = '" & CCGIOTIP & "'"
          Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)

       Else
'  response.write "si "
          perfilsp = "s"

       End If

       If NOT rsPerfil.EOF Then

          If rsPerfil("objsector") Then          bSector = "SI"
            'Sectores

          If rsPerfil("objeven") Then            bEvento = "SI"
            'Eventos

          If rsPerfil("objproced") Then          bProcedimiento = "SI"
            'Procedimientos

          If rsPerfil("objcontrol") Then         bControl = "SI"
            'Controles

          If rsPerfil("objesquema") Then         bEsquema = "SI"
            'Esquemas

          If rsPerfil("objtabla") Then           bTablas = "SI"
            'Tablas

          If rsPerfil("objhtm") Then             bDocum = "SI"
            'Documentos

          If rsPerfil("objobj1") Then            bEditor = "SI"
            'Editor estructurado

          If rsPerfil("objaccion") Then          bDescri = "SI"
            'Descripcion

          If rsPerfil("objmain") Then            bMain = "SI"
            'Consideraciones Generales

        END IF


        Call OFV.cerrarconsulta(rsPerfil)
        Call OFV.cerrarconn(cn)





'''''''''''''''''' invoca generacion del objeto (BLKOBJ)

        xxxx = CopyGenFn("BLKOBJ","NO")


'''''''''''''''''' invoca generacion del relaciones (BLKREL)

        xxxx = CopyGenFn("BLKREL","NO")


'''''''''''''''''' invoca generacion propietario del docuemnto (BLKAUT)

        xxxx = CopyGenFn("BLKAUT","NO")


'''''''''''''''''' invoca generacion propiedades del docuemnto (BLKPRT)

        xxxx = CopyGenFn("BLKPRT","NO")

'''''''''''''''''' invoca generacion perfil del docuemnto (BLKPRF)

        xxxx = CopyGenFn("BLKPRF","NO")


'''''''''''''''''' invoca generacion agrupamiento y segmento (BLKAGS)

        xxxx = CopyGenFn("BLKAGS","NO")

'''''''''''''''''' invoca generacion origen de normativa (BLKORG)

        xxxx = CopyGenFn("BLKORG","NO")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' seccion de bloques segun perfil '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DESCRIPCION ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If bDescri = "SI" Then

'''''''''''''''''' invoca generacion bloque descripcion (BLKDSC)

           xxxx = CopyGenFn("BLKDSC","NO")
        End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' SECTORES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Sectores
        If bSector = "SI" Then
'''''''''''''''''' invoca generacion bloque sectores (BLKSEC)

           xxxx = CopyGenFn("BLKSEC","NO")
        End If



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' EVENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If bEvento = "SI" Then
'''''''''''''''''' invoca generacion bloque eventos (BLKEVE)

           xxxx = CopyGenFn("BLKEVE","NO")
        End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' PROCEDIMIENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      If bProcedimiento = "SI" Then

'''''''''''''''''' invoca generacion bloque procedimientos (BLKPCD)

           xxxx = CopyGenFn("BLKPCD","NO")


      End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DOCUMENTOS ASOCIADOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If bDocum = "SI" Then
'''''''''''''''''' invoca generacion bloque docuemntos asociados (BLKDOC)

           xxxx = CopyGenFn("BLKDOC","NO")
      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' TABLAS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bTablas = "SI" Then
'''''''''''''''''' invoca generacion bloque tablas (BLKTAB)

           xxxx = CopyGenFn("BLKTAB","NO")

      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' EDITOR ESTRUCTURADO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bEditor = "SI" Then
'''''''''''''''''' invoca generacion bloque contenido (BLKEST)

           xxxx = CopyGenFn("BLKEST","NO")

      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' CONTROLES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bControl = "SI" Then
'''''''''''''''''' invoca generacion bloque descripcion (BLKCNT)

           xxxx = CopyGenFn("BLKCNT","NO")
      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' ESQUEMAS CONTABLES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bEsquema = "SI" Then
'''''''''''''''''' invoca generacion bloque descripcion (BLKESQ)

           xxxx = CopyGenFn("BLKESQ","NO")
      End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DOCUMENTOS REDIRECCIONADOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''' invoca generacion bloque descripcion (BLKFAQ)

           xxxx = CopyGenFn("BLKFAQ","NO")



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' CONSIDERACIONES GENERALES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If bMain = "SI" Then

'''''''''''''''''' invoca generacion bloque descripcion (BLKCON)

           xxxx = CopyGenFn("BLKCON","NO")


    End If

 END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FUNCIONES ESPECIALES PARA PROCESOS Y SUBPROCESOS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set cn = OFV.conectar(OFV.strconn2)

''  CCGIOCOD = nCodigo
''  CCGIOTIP = sTipo
''  CCGINCOD = 0
''  CCGINTIP = sTipobjeto
''  CCGICESQ = nCodESQS
''  CCGITESQ = sTipESQS
''  CCGOCESQ = nCodESQSO
''  CCGOTESQ = sTipESQSO

'' sopcion = "TNP"                      ' GENERA ENTRADA PARA EL NUEVO PERIODO - DESDE CERO
'' SOPCION = "TNPDA"                    ' GENERA NUEVO PERIODO - TOMA CIRCUITO DESDE PERIODO ANTERIOR
'' SOPCION = "TNPDAV"                   ' GENERA NUEVO PERIODO - TOMA CIRCUITO DESDE PERIODO ANTERIOR DE LA VERSION ANTERIOR DEL DOC.VINCULADO



 if sopcion = "TNPPL" OR sopcion = "TNP" OR SOPCION = "TNPDA" OR SOPCION = "TNPDAV" THEN


        Strsql = "SELECT OBJID,codobjetoL, tipobjetoL From OBJLNK where CODOBJETO = '" & CCGICESQ & "' and TIPOBJETO = '" & CCGITESQ & "' "
        strsql = strsql & " AND (LTipo = 'LRGVN' OR LTipo = 'LRGCR' ) "

   ''     response.write strsql & "<br>" & "<br>" & "<br>"
        set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsVINC.eof then
           wlnkid = rsVINC(0)
           CCGINCOD = rsVINC(1)
           CCGINTIP = rsVINC(2)

        END IF




    IF SOPCION = "TNPDA" THEN




        WFCCSEC = 0

        Strsql = "SELECT * From ORACTIVRIESGO where codobjeto  = '" & CCGICESQ & "' and tipobjeto = '" &  CCGITESQ & "' AND PERIODO = '" & SPERIODO & "' "
   ''       response.write strsql & "<br>" & chr(13)
        set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsVINC.eof then
           DO WHILE NOT rsVINC.eof

           WFCCSEC = WFCCSEC + 1

           WFCCORIESGO = rsVINC(7)
           WFCCNRIESGO = GTkey(CCGICESQ) & "C" & WFCCSEC


         ''  strsql = "Insert into ORACTIVRIESGO "
         ''  strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
         ''  strsql = strsql & " SECUEN, CODRIESGO, ITEMRIESGO, PROBABILIDAD, IMPACTO, "
         ''  strsql = strsql & " RIESGOBRUTO, RIESGORESIDUAL, VULNERABILIDAD,REQUIEREPLAN,"
         ''  strsql = strsql & " OBSERVACIONES, RIESGORESIDUALNV,REQUIEREINDICADOR"
         ''  strsql = strsql & ") "
         ''  strsql = strsql & " Values ("
         ''  strsql = strsql &  "'" & session("ORPERIODO") & "','" & rsVINC(2) & "','" & rsVINC(3) & "',"
         ''  strsql = strsql &  "'" & rsVINC(4) & "','" & rsVINC(5) & "',"
         ''  strsql = strsql & "'" & rsVINC(6) & "','" & WFCCNRIESGO & "',"
         ''  strsql = strsql & "'" & rsVINC(8) & "','" & rsVINC(9) & "','" & rsVINC(10) & "',"
         ''  strsql = strsql & "'" & rsVINC(11) & "','" & rsVINC(12) & "','" & rsVINC(13) & "','NO',"
         ''  strsql = strsql & "'SIN OBSERVACIONES.','" & rsVINC(17) & "'"
         ''  strsql = strsql &",'" & rsVINC(20) &"')"
      ''    response.write strsql & "<br>" & chr(13)
         ''  call ofv.crearconsultaEx(StrSql,cn,1,volver)



                  strsql = "Insert into ORACTIVRIESGO (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,SECUEN, CODRIESGO"
                  For ii = 8 To (rsVINC.fields.Count-1)
                      strsql = strsql & "," & rsVINC(ii).Name
                  Next
                  strsql = strsql & ") values ('" & session("ORPERIODO") & "','" & rsVINC(2) & "','" & rsVINC(3) & "',"
                  strsql = strsql &  "'" & rsVINC(4) & "','" & rsVINC(5) & "',"
                  strsql = strsql & "'" & rsVINC(6) & "','" & WFCCNRIESGO & "'"


                     strsql2 = ""
                     EY = EY + 1
                     For ii = 8 To (rsVINC.fields.Count-1)
                        sInput = rsVINC(ii)
                        If rsVINC(ii).Type = 131 Then sInput = Replace(rsVINC(ii), ",", ".")

                        If UCase(sInput) = "VERDADERO" Then sInput = "1"
                        If UCase(sInput) = "FALSO" Then sInput = "0"
                        If UCase(sInput) = "TRUE" Then sInput = "1"
                        If UCase(sInput) = "FALSE" Then sInput = "0"


                        strsql2 = strsql2 & ",'" & sInput & "'"
                     Next



            ''    response.write strsql & strsql2 & ")"  & "<br>" & "<br>" & "<br>"

           call ofv.crearconsultaEx(strsql & strsql2 & ")" ,cn,1,volver)





           Strsql = "SELECT * From ORACTIVRCONTROLES where codobjeto  = '" & CCGICESQ & "' and tipobjeto = '" &  CCGITESQ & "' "
           strsql = strsql & " AND PERIODO = '" & SPERIODO & "' AND CODRIESGO = '" & WFCCORIESGO & "' "
       ''   response.write strsql & "<br>" & chr(13)
           set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsVINC2.eof then
              DO WHILE NOT rsVINC2.eof



              strsql = "Insert into ORACTIVRCONTROLES "
              strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
              strsql = strsql & " SECUEN, CODRIESGO, ITEMCONTROL, CALIFMITIGANTE,TIPO, "
              strsql = strsql & " MODOCONT, DESCRICONT "
              strsql = strsql & " ) Values ("
              strsql = strsql &  "'" & session("ORPERIODO") & "','" & rsVINC2(2) & "','" & rsVINC2(3) & "',"
              strsql = strsql &  "'" & rsVINC2(4) & "','" & rsVINC2(5) & "',"
              strsql = strsql & "'" & rsVINC2(9) & "','" & WFCCNRIESGO & "',"
              strsql = strsql & "'" & rsVINC2(7) & "','" & rsVINC2(8) & "','" & rsVINC2(10) & "','" & rsVINC2(11) & "','" & rsVINC2(12) & "'"
              strsql = strsql & ")"
       ''   response.write strsql & "<br>" & chr(13)
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


             rsVINC2.MOVENEXT
             LOOP



           END IF


           Strsql = "SELECT * From ORRELAINDIRIESGO where CODRIESGO  = '" & WFCCORIESGO  & "'  "
           set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsVINC2.eof then
              DO WHILE NOT rsVINC2.eof

              strsql = "Insert into ORRELAINDIRIESGO "
              strsql = strsql & " (IDINDICADOR,CODRIESGO,USUARIO,FECHA)"
              strsql = strsql & " Values ("
              strsql = strsql & "'" & rsVINC2(1) & "','" & WFCCNRIESGO & "',"
              strsql = strsql & "'" & rsVINC2(3) & "','" & rsVINC2(4) & "'"
              strsql = strsql & ")"

           ''     response.write strsql & "<br>" & "<br>" & "<br>"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


             rsVINC2.MOVENEXT
             LOOP



           END IF



          rsVINC.MOVENEXT
          LOOP



        END IF








    END IF


    IF SOPCION = "TNPDAV" THEN


        Strsql = "SELECT OBJID,codobjetoL, tipobjetoL From OBJLNK where CODOBJETO = '" & CCGOCESQ & "' and TIPOBJETO = '" & CCGOTESQ & "' "
        strsql = strsql & " AND (LTipo = 'LRGVN' OR LTipo = 'LRGCR' ) "
    '''            response.write strsql & "<br>" & "<br>" & "<br>"
        set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsVINC.eof then
           wlnkid = rsVINC(0)
           CCGIOCOD = rsVINC(1)
           CCGIOTIP = rsVINC(2)

        END IF


      '''     CCGINCOD = rsVINC(1)
      '''     CCGINTIP = rsVINC(2)


        WFCCSEC = 0

        Strsql = "SELECT c.codproced, c.codaccion ,a.* From ORACTIVRIESGO a "
        strsql = strsql & " inner join "

        strsql = strsql & " ( "
        strsql = strsql & " select x.objid, y.codaccion as relant, x.codaccion, x.codproced, x.codobjeto from objaccion x "
        strsql = strsql & " inner join objaccion y on (x.relant = y.objid ) "
        strsql = strsql & " ) "

        strsql = strsql & "  c on ( a.actividad = c.relant  and c.codobjeto = '" & CCGINCOD & "')  "
        strsql = strsql & " where a.codobjeto  = '" & CCGOCESQ & "' and "
        strsql = strsql & " a.tipobjeto = '" &  CCGOTESQ & "' AND a.PERIODO = '" & SPERIODO & "' "
   ''       response.write strsql & "<br>" & "<br>" & "<br>"
        set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsVINC.eof then
           DO WHILE NOT rsVINC.eof

           WFCCSEC = WFCCSEC + 1

           WFCCORIESGO = rsVINC(9)
           WFCCNRIESGO = GTkey(CCGICESQ) & "C" & WFCCSEC


        ''   strsql = "Insert into ORACTIVRIESGO "
        ''   strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
        ''   strsql = strsql & " SECUEN, CODRIESGO, ITEMRIESGO, PROBABILIDAD, IMPACTO, "
        ''   strsql = strsql & " RIESGOBRUTO, RIESGORESIDUAL, VULNERABILIDAD,REQUIEREPLAN,"
        ''   strsql = strsql & " OBSERVACIONES, RIESGORESIDUALNV,REQUIEREINDICADOR"
        ''   strsql = strsql & ") "
        ''   strsql = strsql & " Values ("
        ''   strsql = strsql &  "'" & session("ORPERIODO") & "','" & rsVINC(1) & "','" & rsVINC(0) & "',"
        ''   strsql = strsql &  "'" & CCGICESQ & "','" & CCGITESQ & "',"
        ''   strsql = strsql & "'" & rsVINC(8) & "','" & WFCCNRIESGO & "',"
        ''   strsql = strsql & "'" & rsVINC(10) & "','" & rsVINC(11) & "','" & rsVINC(12) & "',"
        ''   strsql = strsql & "'" & rsVINC(13) & "','" & rsVINC(14) & "','" & rsVINC(15) & "','NO',"
        ''   strsql = strsql & "'SIN OBSERVACIONES.','" & rsVINC(17) & "'"
        ''   strsql = strsql &",'" & rsVINC(20) &"')"


                  strsql = "Insert into ORACTIVRIESGO (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,SECUEN, CODRIESGO"
                  For ii = 10 To (rsVINC.fields.Count-1)
                      strsql = strsql & "," & rsVINC(ii).Name
                  Next
                  strsql = strsql & ") values ('" & session("ORPERIODO") & "','" & rsVINC(1) & "','" & rsVINC(0) & "',"
                  strsql = strsql &  "'" & CCGICESQ & "','" & CCGITESQ & "',"
                  strsql = strsql & "'" & rsVINC(8) & "','" & WFCCNRIESGO & "'"

               ''''   Do Until rsVINC.EOF
                     strsql2 = ""
                     EY = EY + 1
                     For ii = 10 To (rsVINC.fields.Count-1)
                        sInput = rsVINC(ii)
                        If rsVINC(ii).Type = 131 Then sInput = Replace(rsVINC(ii), ",", ".")

                        If UCase(sInput) = "VERDADERO" Then sInput = "1"
                        If UCase(sInput) = "FALSO" Then sInput = "0"
                        If UCase(sInput) = "TRUE" Then sInput = "1"
                        If UCase(sInput) = "FALSE" Then sInput = "0"


                        strsql2 = strsql2 & ",'" & sInput & "'"
                     Next



            ''    response.write strsql & strsql2 & ")"  & "<br>" & "<br>" & "<br>"

           call ofv.crearconsultaEx(strsql & strsql2 & ")" ,cn,1,volver)



           Strsql = "SELECT * From ORACTIVRCONTROLES where codobjeto  = '" & CCGOCESQ  & "' and tipobjeto = '" &  CCGOTESQ & "' "
           strsql = strsql & " AND PERIODO = '" & SPERIODO & "' AND CODRIESGO = '" & WFCCORIESGO & "' "
           set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsVINC2.eof then
              DO WHILE NOT rsVINC2.eof



              strsql = "Insert into ORACTIVRCONTROLES "
              strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
              strsql = strsql & " SECUEN, CODRIESGO, ITEMCONTROL, CALIFMITIGANTE,TIPO, "
              strsql = strsql & " MODOCONT, DESCRICONT "
              strsql = strsql & " ) Values ("
              strsql = strsql &  "'" & session("ORPERIODO") & "','" & rsVINC(1) & "','" & rsVINC(0) & "',"
              strsql = strsql &  "'" & CCGICESQ & "','" & CCGITESQ & "',"
              strsql = strsql & "'" & rsVINC2(9) & "','" & WFCCNRIESGO & "',"
              strsql = strsql & "'" & rsVINC2(7) & "','" & rsVINC2(8) & "','" & rsVINC2(10) & "','" & rsVINC2(11) & "','" & rsVINC2(12) & "'"
              strsql = strsql & ")"

           ''     response.write strsql & "<br>" & "<br>" & "<br>"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)






             rsVINC2.MOVENEXT
             LOOP



           END IF


           Strsql = "SELECT * From ORRELAINDIRIESGO where CODRIESGO  = '" & WFCCORIESGO  & "'  "
   ''             response.write strsql & "<br>" & "<br>" & "<br>"
           set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsVINC2.eof then
              DO WHILE NOT rsVINC2.eof

              strsql = "Insert into ORRELAINDIRIESGO "
              strsql = strsql & " (IDINDICADOR,CODRIESGO,USUARIO,FECHA)"
              strsql = strsql & " Values ("
              strsql = strsql & "'" & rsVINC2(1) & "','" & WFCCNRIESGO & "',"
              strsql = strsql & "'" & rsVINC2(3) & "','" & rsVINC2(4) & "'"
              strsql = strsql & ")"

        ''        response.write strsql & "<br>" & "<br>" & "<br>"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


             rsVINC2.MOVENEXT
             LOOP



           END IF



          rsVINC.MOVENEXT
          LOOP



        END IF







    END IF

''  CCGINCOD = 0
''  CCGINTIP = sTipobjeto


''''''''''' GENERO INICIALIZACION DEL CIRCUITO PARA AUTOEVALUACION

              strsql = "Insert into ORRSKCIRCUITO "
              strsql = strsql & " (PERIODO,codobjeto, tipobjeto,"
              strsql = strsql & " OWNER, FECCRE, FECVIGEN, FECBAJA,"
              strsql = strsql & " ESTADO, APLIC, CODOBJETOBS, TIPOBJETOBS,"
              strsql = strsql & " SEGCT, SEGMO, SEGCO, SEGCE"
              strsql = strsql & ") "
              strsql = strsql & " Values ("
              strsql = strsql &  "'" & session("ORPERIODO") & "','" & CCGICESQ & "','" & CCGITESQ & "',"
              strsql = strsql &  "'" & SESSION("USUARIO") & "','" & (CLNG(DATE()) - CLNG(CDATE("01-01-1900"))) & "','0','0',"
              strsql = strsql & "'1','" & wkaplic & "','" & CCGINCOD & "','" & CCGINTIP & "',"
              strsql = strsql & "'0','0','0','0'"
              strsql = strsql & ")"
         ''       response.write strsql & "<br>" & "<br>" & "<br>"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


       IF SOPCION = "TNPPL" THEN

'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

              WKSPTIPO       = "PLOBJ"
              WKSPNOMBRE     = CCGICESQ & "-" & CCGITESQ & "- SOPORTE-CIRCUITO-CONTINGENCIAS - " & session("ORPERIODO")
              WKSPCODIDEN    = CCGICESQ & "-" & CCGITESQ & "- PL - " & session("ORPERIODO")
              WKSPAPLICSET   = "PL"
              WKSPTIPOOBJ    = "PLSPCIRC"

              strsql = "Insert into OBJETOS "
              strsql = strsql & " (tipobjeto, DESCRI, OBJUN, APLIC, "
              strsql = strsql & " CODVINOBJ, RELLINK, PRIORIDAD, CODIDEN,"
              strsql = strsql & " OBJREM, OBJRDO, OBJNIVEL, TSEGMENTO,"
              strsql = strsql & " TAGRUPA"
              strsql = strsql & ") "
              strsql = strsql & " Values ("
              strsql = strsql &  "'" & WKSPTIPOOBJ & "','" & WKSPNOMBRE & "','" & SESSION("UN") & "','" & WKSPAPLICSET & "',"
              strsql = strsql &  "'" & WKSPTIPO & "','0','1','" & cstr(session.sessionid) & "',"
              strsql = strsql & "'0','0','9','0',"
              strsql = strsql & "'0'"
              strsql = strsql & ")"
'''''' response.write "primer insert " & strsql & "<br>" & chr(13)
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


              if cn.errors.count = 0 then

                strsql = "Select codobjeto From Objetos Where codiden = '" & cstr(session.sessionid) & "'"

                set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
                if not rs.eof then nCodAsoc = rs(0)

                call ofv.cerrarconsulta(rs)

                ncodiden =  WKSPCODIDEN


                strsql = "Update Objetos set codiden = '" & ncodiden & "', indobj = 10 "
                strsql = strsql & " where codiden = '" & cstr(session.sessionid) & "'"
                call ofv.crearconsultaEx(StrSql,cn,1,volver)


                Strsql = "insert into OBJdistPrty  (codobjeto,tipobjeto,emisor,asunto,texto,importancia,aviso,dias) values ("
                Strsql = Strsql & nCodAsoc & ",'" & WKSPTIPOOBJ & "','" & session("usuario") & "','Activacion de <ESQNOR>',"
                Strsql = Strsql & "'<ESQNOR> esta vigente a partir de la fecha',1,'NO','7')"
                call ofv.crearconsultaEx(StrSql,cn,1,volver)

' response.write "1-2do insert " & strsql & "<br>" & chr(13)


                IF SESSION("EMPMUN") = "S" THEN

                   Strsql = "insert into OBJdist  (codobjeto,tipobjeto,GRPUSR,SECTOR) values ("
                   Strsql = Strsql & nCodAsoc & ",'" & WKSPTIPOOBJ & "','TODOS','" & session("UN") & "')"
                   call ofv.crearconsultaEx(StrSql,cn,1,volver)

                END IF

                strsql = "Insert into Relaobj (Codobjmaster,Tipobjmaster,Codojasoc,Tipobjasoc) "
                strsql = strsql & "Values (" & CCGICESQ & ",'" & CCGITESQ & "',"
                strsql = strsql & nCodAsoc & ",'" & WKSPTIPOOBJ & "')"
' response.write "2do insert " & strsql & "<br>" & chr(13)
                call ofv.crearconsultaEx(StrSql,cn,1,volver)

                nEstado = 1

                dFechaNula = 0

                dHoy = clng(Date) - clng(cdate("01-01-1900"))

                strsql = "Insert into Objprty (codobjeto, tipobjeto, estado, fecemis, fecvigen, "
                strsql = strsql & "fecbaja, fecvto, fecacta,fecpublic, autor) Values (" & nCodAsoc & ",'" & WKSPTIPOOBJ & "',"
                strsql = strsql & nEstado & ",'" & DHoy & "','" & DFechaNula & "','"
                strsql = strsql & DFechaNula & "','" & DFechaNula & "','" & dFechaNula & "','" & dFechaNula & "','" & session("usuario") & "')"

' response.write "3er insert " & strsql & "<br>" & chr(13)
                call ofv.crearconsultaEx(StrSql,cn,1,volver)

                if cn.errors.count = 0 then
                   strsql = "Insert into OBJSDK (codobjeto, tipobjeto, segct,segmo,segco,segce,segown "
                   strsql = strsql & ") Values (" & nCodAsoc & ",'" & WKSPTIPOOBJ & "',0,0,0,0,'"
                   strsql = strsql & Session("Usuario") & "')"
                   call ofv.crearconsultaEx(StrSql,cn,1,volver)

                   if cn.errors.count = 0 then
                      strsql = "Insert into OBJWFL (codobjeto, tipobjeto)"
                      strsql = strsql & " Values (" & nCodAsoc & ",'" & WKSPTIPOOBJ & "')"
' response.write "4to insert " & strsql & "<br>" & chr(13)
                      call ofv.crearconsultaEx(StrSql,cn,1,volver)

                      if cn.errors.count = 0 then


                              GETFECH()

                              WKFECHA = XLFECDAT - clng(cdate("01-01-1900"))
                              WKHORA  = XLHORDAT
                              WKSECTOR = 0


                              strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, iden3, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
                              strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values ('" & CCGICESQ & "','" & CCGITESQ & "',"
                              strsql = strsql & "'" & CCGICESQ & "','" & CCGICESQ & "','" & SESSION("ORPERIODO") & "',"
                              strsql = strsql & "'" & nCodAsoc & "','" & WKSPTIPOOBJ & "','','','LRGPL','0','0','0'"
                              strsql = strsql & ")"
                              call ofv.crearconsultaEx(StrSql,cn,1,volver)

                              Strsql = "SELECT * From OBJLNK where iden1 = '" & CCGICESQ & "' and "
                              strsql = strsql & " iden2 = '" & CCGICESQ & "' and "
                              strsql = strsql & " iden3 = '" & SESSION("ORPERIODO") & "' and "
                              strsql = strsql & "LTIPO = 'LRGPL' "
                              set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                              if not rsN.eof then
                                 wlnkid = rsN(0)
                              else
                                 wlnkid = 0
                                 wlink = 0
                              end if
                              call ofv.cerrarconsulta(rsN)

                              if clng(wlnkid) > 0 then
                                 strsql = "update ORRSKCIRCUITO set rellink = " & wlnkid
                                 strsql = strsql & " where  PERIODO = '" & SESSION("ORPERIODO") & "' AND "
                                 strsql = strsql & " CODOBJETO = '" & CCGICESQ & "' and "
                                 strsql = strsql & " TIPOBJETO = '" & CCGITESQ & "' AND "
                                 strsql = strsql & " APLIC = 'PL'  "
                                call ofv.crearconsultaEx(StrSql,cn,1,volver)
                              end if



                      END IF
                   END IF
                END IF
             END IF



'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

       END IF


 ELSE


       IF SOPCION = "RGCRNV" THEN

          CCGICESQ = CCGINCOD
          CCGITESQ = CCGINTIP

          CCGOCESQ = CCGIOCOD
          CCGOTESQ = CCGIOTIP

          wlnkid   = 0

          Strsql = "SELECT OBJID,codobjetoL, tipobjetoL,IDEN4, LTIPO From OBJLNK where iden1 = '" & CCGOCESQ & "' and iden2 = '" & CCGOCESQ & "' "
          strsql = strsql & " AND (LTipo = 'LRGVN' OR LTipo = 'LRGCR' ) "
          set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsVINC.eof then
             wlnkid = rsVINC(0)
             wlnkidX1 = rsVINC(1)
             wlnkidX2 = rsVINC(2)
             wlnkidX3 = rsVINC(3)
             wlnkTIPL = rsVINC(4)

             CCGIOCOD = wlnkidX1
             CCGIOTIP = wlnkidX2
             CCGINTIP = wlnkidX2

          end if

          if wlnkid   > 0 then


          '''''''''''''''''' invoca generacion del objeto (BLKOBJ)

             xxxx = CopyGenFn("BLKOBJ","NO")


          '''''''''''''''''' invoca generacion del relaciones (BLKREL)

             xxxx = CopyGenFn("BLKREL","NO")


          '''''''''''''''''' invoca generacion propietario del docuemnto (BLKAUT)

             xxxx = CopyGenFn("BLKAUT","NO")


          '''''''''''''''''' invoca generacion propiedades del docuemnto (BLKPRT)

             xxxx = CopyGenFn("BLKPRT","NO")

          '''''''''''''''''' invoca generacion perfil del docuemnto (BLKPRF)

             xxxx = CopyGenFn("BLKPRF","NO")


          '''''''''''''''''' invoca generacion agrupamiento y segmento (BLKAGS)

             xxxx = CopyGenFn("BLKAGS","NO")

          '''''''''''''''''' invoca generacion origen de normativa (BLKORG)

             xxxx = CopyGenFn("BLKORG","NO")

          '''''''''''''''''' invoca generacion bloque sectores (BLKSEC)

             xxxx = CopyGenFn("BLKSEC","NO")

          '''''''''''''''''' invoca generacion bloque eventos (BLKEVE)

             xxxx = CopyGenFn("BLKEVE","NO")

          '''''''''''''''''' invoca generacion bloque procedimientos (BLKPCD)

             xxxx = CopyGenFn("BLKPCD","NO")





              strsql = "SELECT TIPOBJETO,CODPROCED FROM OBJPROCED    "
              strsql = strsql & " WHERE CODOBJETO = 0" & CCGINCOD & " AND RELANT = '" & wlnkidX3 & "' "
       '''   response.write strsql & "<br>"
              set RSOBJOLD = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              if NOT RSOBJOLD.eof then

                 NDOCUTIPO  = RSOBJOLD(0)
                 NPROCEDCOD = RSOBJOLD(1)

              END IF


                   wlink   = CCGINCOD
                   wlinkT  = CCGINTIP
                   wlinkN1 = "5"
                   wlinkN2 = NPROCEDCOD
                   wlinkN3 = "0"

                   wlinkDSC = CCGICESQ & "-" & CCGITESQ & "- VINCULACION EXPLICITA CR"



                      strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
                      strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values (" & CCGICESQ & ",'" & CCGITESQ & "',"
                      strsql = strsql & "'" & CCGICESQ & "','" & CCGICESQ & "',"
                      strsql = strsql & wlink & ",'" & wlinkT & "','','','" & wlnkTIPL & "','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
                      strsql = strsql & ")"
                      call ofv.crearconsultaEx(StrSql,cn,1,volver)

                      Strsql = "SELECT * From OBJLNK where iden1 = '" & CCGICESQ & "' and iden2 = '" & CCGICESQ & "' AND LTipo = '" & wlnkTIPL & "' "
                      set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rsVINC2.eof then
                         wlnkid = rsVINC2(0)

                         strsql = "update OBJETOS set rellink = " & wlnkid & ", APLIC='RG' "
                         strsql = strsql & " where  codobjeto = " & CCGICESQ
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJETOS set  APLIC='RG', DESCRI = '" & wlinkDSC & "' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 3, VERSION = 'V', NROVER = '1' "
                         strsql = strsql & " where  codobjeto = " & CCGICESQ
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 3, VERSION = 'V', NROVER = '1' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 4, FECBAJA = '" & (CLNG(DATE()) - CLNG(CDATE("01-01-1900"))) & "' "
                         strsql = strsql & " where  codobjeto = " & CCGOCESQ
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 4, FECBAJA = '" & (CLNG(DATE()) - CLNG(CDATE("01-01-1900"))) & "' "
                         strsql = strsql & " where  codobjeto = " & CCGIOCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                      end if
                      call ofv.cerrarconsulta(rsVINC2)

           else


              Response.write "**!! Fallo reemplazo nueva version del subproceso: " & Minute(Now) & ":" & Second(Now) & "<br>"



           end if




       else


        Strsql = "SELECT OBJID,codobjetoL, tipobjetoL,IDEN4, LTIPO From OBJLNK where iden1 = '" & CCGIOCOD & "' and iden2 = '" & CCGIOCOD & "' "
        strsql = strsql & " AND (LTipo = 'LRGVN' OR LTipo = 'LRGCR' ) "
        set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsVINC.eof then
           wlnkid = rsVINC(0)
           wlnkidX1 = rsVINC(1)
           wlnkidX2 = rsVINC(2)
           wlnkidX3 = rsVINC(3)
           wlnkTIPL = rsVINC(4)



           strsql = "SELECT OBJRDO FROM OBJETOS    "
           strsql = strsql & " WHERE CODOBJETO = 0" & wlnkidX1
       '''   response.write strsql & "<br>"
           set RSOBJOLD = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if NOT RSOBJOLD.eof then

              NDOCUCOD = RSOBJOLD(0)
              NDOCUTIPO = RSOBJOLD(0)
              NPROCEDCOD = RSOBJOLD(0)

              strsql = "SELECT TIPOBJETO,CODPROCED FROM OBJPROCED    "
              strsql = strsql & " WHERE CODOBJETO = 0" & NDOCUCOD & " AND RELANT = '" & wlnkidX3 & "' "
       '''   response.write strsql & "<br>"
              set RSOBJOLD = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              if NOT RSOBJOLD.eof then

                 NDOCUTIPO  = RSOBJOLD(0)
                 NPROCEDCOD = RSOBJOLD(1)

              END IF


                   wlink   = NDOCUCOD
                   wlinkT  = NDOCUTIPO
                   wlinkN1 = "5"
                   wlinkN2 = NPROCEDCOD
                   wlinkN3 = "0"




                      strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
                      strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
                      strsql = strsql & "'" & CCGINCOD & "','" & CCGINCOD & "',"
                      strsql = strsql & wlink & ",'" & wlinkT & "','','','" & wlnkTIPL & "','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
                      strsql = strsql & ")"
                      call ofv.crearconsultaEx(StrSql,cn,1,volver)

                      Strsql = "SELECT * From OBJLNK where iden1 = '" & CCGINCOD & "' and iden2 = '" & CCGINCOD & "' AND LTipo = '" & wlnkTIPL & "' "
                      set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rsVINC2.eof then
                         wlnkid = rsVINC2(0)

                         strsql = "update OBJETOS set rellink = " & wlnkid & ", "
                         strsql = strsql & " OBJREM = " & CCGIOCOD & " "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 3, VERSION = 'V', NROVER = '1' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJETOS set OBJRDO = " & CCGINCOD
                         strsql = strsql & " where  codobjeto = " & CCGIOCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 4, FECBAJA = '" & (CLNG(DATE()) - CLNG(CDATE("01-01-1900"))) & "' "
                         strsql = strsql & " where  codobjeto = " & CCGIOCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                      end if
                      call ofv.cerrarconsulta(rsVINC2)



          END IF


       end if
        call ofv.cerrarconsulta(rsVINC)



    end if


   END IF


        Call OFV.cerrarconn(cn)




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Response.write "tiempo  fin: " & Minute(Now) & ":" & Second(Now) & "<br>"


        Response.write "<center>"

        sInput = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "")
        Response.write "<form name=volver action='" & volver & "' method='post' >" & sInput & "</form>"
        Response.write "</center>"

        sHTML = "<SCRIPT LANGUAGE='JAVASCRIPT' >window.top.ldfobj();</SCRIPT></body></html>"


       ' response.redirect volver








end if

response.write sHTML
'response.redirect volver

%>