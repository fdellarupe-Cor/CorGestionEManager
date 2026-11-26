<%

 dim CCGIOCOD
 dim CCGIOTIP
 dim CCGINCOD
 dim CCGINTIP
 dim CCGIEMPR
 dim CCGINPRF
 dim CCGIDESC
 dim CCGICDID
 dim CCGICESQ
 dim CCGITESQ
 dim CCGIXLNK
 dim CCGIXORG


 dim CCGminini
 dim CCGsecini
 dim CCGminfin
 dim CCGsecfin

 ReDim Preserve asector(0)
 ReDim Preserve nsector(0)
 ReDim Preserve aevento(0)
 ReDim Preserve nevento(0)
 ReDim Preserve ProcOld(0)
 ReDim Preserve ProcNew(0)
 ReDim Preserve AccPnew(0)
 ReDim Preserve AccIOld(0)
 ReDim Preserve AccINew(0)
 ReDim Preserve aestr(0)
 ReDim Preserve nestr(0)
 ReDim Preserve aestdet(0)
 ReDim Preserve nestdet(0)
 ReDim Preserve ycontrol(0)
 ReDim Preserve Ocontrol(0)
 ReDim Preserve Oesque(0)
 ReDim Preserve Nesque(0)
 ReDim Preserve aTabla(0)
 ReDim Preserve nTabla(0)
 ReDim Preserve aTablab(0)
 ReDim Preserve nTablab(0)

Function CopyGenFn(CCGIBLK,CCGIOPC)

'''''''''''''''''''''''''''''''''''''''



        nFecini = 0
        If nFecini < CLng(CDate("01-01-1900")) Then
            nFecini = CLng(CDate("01-01-1900"))
        Else
            nFecini = 0
        End If

        dFechaNula = CLng(CDate("01-01-1900")) - nFecini
        Nfechoy    = CLng(Date) - nFecini
        sUsuario   = Session("Usuario")


''''''''''''''''''''''''''''''''''''''''''''''''''''


    select case CCGIBLK

           case "BLKOBJ"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DOCUMENTO PRINCIPAL (OBJETO) ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                 Set CCGICN = OFV.conectar(OFV.strconn2)


                 strsql    = "Select objnivel From Objetos Where codobjeto = " & CCGIOCOD
                 Set rsNew = OFV.crearconsulta2(strsql, CCGICN)
                 If Not rsNew.EOF Then
                     nOBJNIV = rsNew(0)
                 Else
                     nOBJNIV = 9
                 End If
                 Call OFV.cerrarconsulta(rsNew)

                 wsval1 = CCGIDESC
                 strsql = "Select * From Objetos Where codobjeto = " & CCGIOCOD
                 Set rsNew = OFV.crearconsulta2(strsql, CCGICN)
                 If Not rsNew.EOF Then
                    If IsNull(rsNew(8)) Then
                       wsval1 = CCGIDESC
                    Else
                       wsval1 = rsNew(8)
                    End If
                 End If



                 strsql = "Insert Into Objetos (codiden,tipobjeto,descri,objrem,objrdo,objnivel,descriabr,"
                 strsql = strsql & "TSEGMENTO,TAGRUPA,CODIDEN2,CODIDEN3,OBJUN,APLIC,"
                 strsql = strsql & "INDOBJ,RELLINK,REFLINK,CODVINOBJ,LINNEG,PRIORIDAD"
                 strsql = strsql & ") "
                 strsql = strsql & "Values ('" & CCGICDID & "','" & CCGINTIP & "','" & Session.sessionid & "',"
                 strsql = strsql & CCGIOCOD & ",0," & nOBJNIV & ",'" & wsval1 & "','0" & rsNew(9) & "'"
                 strsql = strsql & ",'0" & rsNew(10) & "','" & rsNew(11) & "','" & rsNew(12) & "','"
                 strsql = strsql & CCGIEMPR & "','" & wkaplic & "',"
                 strsql = strsql & "'0" & rsNew(13) & "','0" & rsNew(15) & "','" & rsNew(16) & "',"
                 strsql = strsql & "'" & rsNew(17) & "','0" & rsNew(18) & "','0" & rsNew(19) & "'"
                 strsql = strsql & ")"
                 Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

                 Call OFV.cerrarconsulta(rsNew)

                 strsql = "Select codobjeto From Objetos Where descri = '" & Session.sessionid & "'"
                 Set rsNew = OFV.crearconsulta2(strsql, CCGICN)
                 If Not rsNew.EOF Then
                     rsNew.movelast
                     CCGINCOD = rsNew(0)
                   '  Response.write "CCGINCOD=" & CCGINCOD & "<br>"
                 End If
                 Call OFV.cerrarconsulta(rsNew)

                 strsql = "Update Objetos set descri = '" & CCGIDESC & "' Where descri = '" & Session.sessionid & "'"
                 Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 Call OFV.cerrarconn(CCGICN)


           case "BLKREL"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Relaciones DEL OBJETO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                 Set CCGICN = OFV.conectar(OFV.strconn2)

                 strsql = "Insert into Relaobj (Codobjmaster,Tipobjmaster,Codojasoc,Tipobjasoc) "
                 strsql = strsql & "Values (" & CCGICESQ & ",'" & CCGITESQ & "'," & CCGINCOD & ",'"
                 strsql = strsql & CCGINTIP & "')"
                 Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 Call OFV.cerrarconn(CCGICN)


           case "BLKAUT"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' PROPIETARIO DEL OBJETO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                 Set CCGICN = OFV.conectar(OFV.strconn2)


				   Strsql = "insert into OBJdistPrty  (codobjeto,tipobjeto,emisor,asunto,texto,importancia,aviso,dias) values ("
				   Strsql = Strsql & CCGINCOD & ",'" & CCGINTIP & "','" & session("usuario") & "','Activacion de <ESQNOR>',"
				   Strsql = Strsql & "'<ESQNOR> esta vigente a partir de la fecha',1,'NO','7')"
				   call ofv.crearconsultaEx(StrSql,cn,1,volver)



                 strsql = "Insert into OBJSDK (codobjeto, tipobjeto, segct,segmo,segco,segce,segown "
                 strsql = strsql & ") Values (" & CCGINCOD & ",'" & CCGINTIP & "',0,0,0,0,'"
                 strsql = strsql & Session("Usuario") & "')"

                 Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 Call OFV.cerrarconn(CCGICN)


           case "BLKPRT"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' PROPIEDADES DEL OBJETO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                 Set CCGICN = OFV.conectar(OFV.strconn2)
                 strsql = "Select * From Objprty Where codobjeto = " & CCGIOCOD
                 Set rsPrty = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

                 strsql = "Insert Into Objprty (codobjeto,tipobjeto,origen,instrumen,instid,estado,"
                 strsql = strsql & "autor,fecemis,fecvigen,fecbaja,fecvto,fecacta,fecpublic,version,nrover,prtylink,prtylink2)"
                 strsql = strsql & " Values (" & CCGINCOD & ",'"
                 strsql = strsql & CCGINTIP & "'," & rsPrty(3) & "," & rsPrty(4) & ",'" & rsPrty(5)
                 strsql = strsql & "',1,'" & sUsuario & "','" & Nfechoy & "','" & dFechaNula & "','"
                 strsql = strsql & dFechaNula & "','" & dFechaNula & "','" & dFechaNula & "','" & dFechaNula & "','"
                 strsql = strsql & rsPrty(14) & "','" & rsPrty(15) & "','" & rsPrty(18) & "','" & rsPrty(19) & "')"

            ''     Response.write "<br>" & strsql & "<br>"

                 Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 Call OFV.cerrarconsulta(rsPrty)
                 Call OFV.cerrarconn(CCGICN)


           case "BLKPRF"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' PERFIL DEL DOCUMENTO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                 Set CCGICN = OFV.conectar(OFV.strconn2)


                 strsql = "Select * From Objperfildoc Where codobjeto =" & CCGIOCOD
                 Set rCCGINPRF = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 If NOT rCCGINPRF.EOF Then

                    strsql = "Insert into Objperfildoc (codobjeto,tipobjeto"
                    strsql2 = ") values (" & CCGINCOD & ",'" & CCGINTIP & "'"

                    For ii = 0 To (rCCGINPRF.fields.Count - 4)
                        strsql = strsql & "," & rCCGINPRF(ii).Name
                        s = OFV.QueEsx(rCCGINPRF(ii).Type)
                        If rCCGINPRF(ii).Type > 100 And rCCGINPRF(ii).Type < 300 Then s = "'"

                        sInput = rCCGINPRF(ii)

                        If UCase(sInput) = "VERDADERO" Then sInput = "1"
                        If UCase(sInput) = "FALSO" Then sInput = "0"
                        If UCase(sInput) = "TRUE" Then sInput = "1"
                        If UCase(sInput) = "FALSE" Then sInput = "0"

                        If rCCGINPRF(ii).Type = 131 Then sInput = Replace(rCCGINPRF(ii), ",", ".")

                        strsql2 = strsql2 & "," & s & sInput & s

                    Next

                    strsql = strsql & strsql2 & ") "
                    Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 End If

                 Call OFV.cerrarconsulta(rCCGINPRF)
                 Call OFV.cerrarconn(CCGICN)



           case "BLKAGS"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' AGRUPAMIENTO Y SEGMENTO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Set CCGICN = OFV.conectar(OFV.strconn2)
                strsql = "Select * From ObjTSEGAGR Where codobjeto = " & CCGIOCOD
                Set rshtm = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                if not rshtm.eof then
                   EY = 0

                  strsql = "Insert into ObjTSEGAGR (codobjeto,tipobjeto"
                  For ii = 3 To (rshtm.fields.Count-1)
                      strsql = strsql & "," & rshtm(ii).Name
                  Next
                  strsql = strsql & ") values (" & CCGINCOD & ",'" & CCGINTIP & "'"


                  Do Until rshtm.EOF
                     strsql2 = ""
                     EY = EY + 1
                     For ii = 3 To (rshtm.fields.Count-1)
                        sInput = rshtm(ii)
                        If rshtm(ii).Type = 131 Then sInput = Replace(rshtm(ii), ",", ".")
                        strsql2 = strsql2 & ",'" & sInput & "'"
                     Next

                     Call OFV.crearconsultaEx(strsql & strsql2 & ") ", CCGICN, 1, volver)
                     rshtm.movenext
                  Loop
                  Call OFV.cerrarconsulta(rshtm)
                  Call OFV.cerrarconn(CCGICN)

                end if

           case "BLKORG"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' ORIGEN DE NORMATIVA ADICIONALES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                 Set CCGICN = OFV.conectar(OFV.strconn2)
                 strsql = "Select * From Objprtyadv Where codobjeto = " & CCGIOCOD
                 Set rshtm = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 if not rshtm.eof then
                    EY = 0

                   strsql = "Insert into Objprtyadv (codobjeto,tipobjeto"
                   For ii = 3 To (rshtm.fields.Count-1)
                       strsql = strsql & "," & rshtm(ii).Name
                   Next
                   strsql = strsql & ") values (" & CCGINCOD & ",'" & CCGINTIP & "'"


                   Do Until rshtm.EOF
                      strsql2 = ""
                      EY = EY + 1
                      For ii = 3 To (rshtm.fields.Count-1)
                         sInput = rshtm(ii)
                         If rshtm(ii).Type = 131 Then sInput = Replace(rshtm(ii), ",", ".")
                         strsql2 = strsql2 & ",'" & sInput & "'"
                      Next

                      Call OFV.crearconsultaEx(strsql & strsql2 & ") ", CCGICN, 1, volver)
                      rshtm.movenext
                   Loop
                   Call OFV.cerrarconsulta(rshtm)
                   Call OFV.cerrarconn(CCGICN)

                 end if



           case "BLKDSC"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DESCRIPCION ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                 Set CCGICN = OFV.conectar(OFV.strconn2)
                 strsql = "Select * From Objdscabr Where codobjeto = " & CCGIOCOD
                 Set rsEvento = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                 Do Until rsEvento.EOF
                     strsql = "Insert Into Objdscabr (codobjeto,tipobjeto,descri,titulo,muestra) Values ("
                     strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "','" & rsEvento(3) & "','"
                     strsql = strsql & rsEvento(4) & "','" & rsEvento(5) & "')"
                     Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                     rsEvento.movenext
                 Loop
                 Call OFV.cerrarconsulta(rsEvento)
                 Call OFV.cerrarconn(CCGICN)



           case "BLKSEC"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' SECTORES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                  IF CCGIXORG = "PL" THEN
                     CCGIWRFLK = "PL"
                  ELSE
                     CCGIWRFLK = ""
                  END IF


                  Set CCGICN = OFV.conectar(OFV.strconn2)
                  strsql = "Select * From Objsector Where codobjeto = " & CCGIOCOD & " order by objid"
                  Set rsSector = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                  SY = 0
                  si = 0
                  Do Until rsSector.EOF
                      ReDim Preserve asector(si)
                      asector(si) = rsSector(0)
                      si = si + 1
                      SY = SY + 1
                      strsql = "Insert Into Objsector (codobjeto,tipobjeto,calificador,codsector,"
                      strsql = strsql & "modi,relant,RELLINK,REFLINK) "
                      strsql = strsql & "Values (" & CCGINCOD & ",'" & CCGINTIP & "','" & rsSector(3)
                      strsql = strsql & "'," & rsSector(4) & ",0,'" & rsSector(0) & "',"
                      strsql = strsql & "'0','" & CCGIWRFLK & "')"

                    '  response.write strsql & "<br>"
                      Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                      rsSector.movenext
                  Loop
                  Call OFV.cerrarconsulta(rsSector)

                  strsql = "Select * From Objsector Where codobjeto = " & CCGINCOD & " order by objid"
                  Set rsSector = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                  si = 0
                  Do Until rsSector.EOF
                      ReDim Preserve nsector(si)
                      nsector(si) = rsSector(0)
                      si = si + 1
                      rsSector.movenext
                  Loop
                  Call OFV.cerrarconsulta(rsSector)
                  Call OFV.cerrarconn(CCGICN)



           case "BLKEVE"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' EVENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                  IF CCGIXORG = "PL" THEN
                     CCGIWRFLK = "PL"
                  ELSE
                     CCGIWRFLK = ""
                  END IF

                   Set CCGICN = OFV.conectar(OFV.strconn2)
                   strsql = "Select * From Objeven Where codobjeto = " & CCGIOCOD & " order by objid"
                   Set rsEvento = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                   EY = 0
                   ei = 0
                   Do Until rsEvento.EOF
                       EY = EY + 1
                       ReDim Preserve aevento(ei)
                       aevento(ei) = rsEvento(0)
                       ei = ei + 1
                       strsql = "Insert Into Objeven (codobjeto,tipobjeto,descri,secuen,modi,relant,codproceso,"
                       strsql = strsql & "RELLINK,REFLINK) "
                       strsql = strsql & " Values ("
                       strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "','" & replace(rsEvento(3),"'","''") & "',"
                       strsql = strsql & rsEvento(4) & ",0,'" & rsEvento(0) & "','0" & rsEvento(11) & "',"
                       strsql = strsql & "'0','" & CCGIWRFLK & "')"
                       Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                       rsEvento.movenext
                   Loop
                   Call OFV.cerrarconsulta(rsEvento)

                   strsql = "Select * From Objeven Where codobjeto = " & CCGINCOD & " order by objid"
                   Set rsEvento = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                   ei = 0
                   Do Until rsEvento.EOF
                       ReDim Preserve nevento(ei)
                       nevento(ei) = rsEvento(0)
                       ei = ei + 1
                       rsEvento.movenext
                   Loop
                   Call OFV.cerrarconsulta(rsEvento)
                   Call OFV.cerrarconn(CCGICN)



           case "BLKPCD"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' PROCEDIMIENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                  IF CCGIXORG = "PL" THEN
                     CCGIWRFLK = "PL"
                  ELSE
                     CCGIWRFLK = ""
                  END IF


                  Set CCGICN = OFV.conectar(OFV.strconn2)
                  strsql = "Select * From Objproced Where codobjeto = " & CCGIOCOD & " order by codproced "
                  Set rsProced = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
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
                    ReDim Preserve ProCNew(i)
                    inp = inp + 1
                    nNewProced = rsProced(7)
                    nProced = GTkey(CCGINCOD & inp)
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

                    IF CCGIXORG = "PL" THEN
                       WKWTIPPROCED = "PL"
                    ELSE
                       WKWTIPPROCED = rsProced(18)
                    END IF



                    strsql = "Insert Into Objproced (codobjeto,tipobjeto,codeven,codsector,codperfil,"
                    strsql = strsql & "secuencia,codproced,frecuencia,referencia,modi,relant,"
                    strsql = strsql & "RELLINK,REFLINK,TIPPROCED) "
                    strsql = strsql & " Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
                    strsql = strsql & Neven & "," & Nsect & ",'" & rsProced(5) & "',"
                    strsql = strsql & rsProced(6) & ",'" & nProced & "','" & rsProced(8) & "','"
                    strsql = strsql & rsProced(9) & "',0,'" & rsProced(7) & "',"
                    strsql = strsql & "'0','" & CCGIWRFLK & "','" & WKWTIPPROCED & "')"
                    Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                    rsProced.movenext
                    i = i + 1
                    ip = ip + 1
                  Loop
                  Call OFV.cerrarconsulta(rsProced)
                  Call OFV.cerrarconn(CCGICN)
                  Response.write "<p> Procedimientos: " & inp & " </p>"
            '-----------------------------------------------
            '  Procedimientos - Acciones
                  Set CCGICN = OFV.conectar(OFV.strconn2)
                  strsql = "Select * From Objaccion Where codobjeto = " & CCGIOCOD
                  strsql = strsql & " and (OBJACCION.modi = 0 or (OBJACCION.modi <> 0 "
                  strsql = strsql & " and OBJACCION.tipmodi <> 'E'))  order by codproced , codaccion"
                  Set rs1 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

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
                     sukey = GTkey(CCGINCOD & ind)

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

                    IF CCGIXORG = "PL" THEN
                       WKWTIPPROCED = "PL"
                    ELSE
                       WKWTIPPROCED = rs1(29)
                    END IF

                     wskllnk = "0" & cstr(rs1(6) & "")

                 If clng(rs1(6)) > 0 and CCGIXLNK <> "N" Then
                    strsql = "SELECT * From OBJLNK where objid = 0" & rs1(6)
                    Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                    If Not rsDsc.EOF Then

                        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
                        strsql = strsql & ",iden3l,iden4,iden4l,linkasoc,postxt,canttxt,dattxt) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
                        strsql = strsql & "'" & nProced & "','" & sukey & "','" & rsDsc("iden3") & "',"
                        strsql = strsql & rsDsc("codobjetoL") & ",'" & rsDsc("tipobjetoL") & "','" & rsDsc("LTipo") & "','" & rsDsc("Lurl") & "','" & rsDsc("LurlDesc") & "'"
                        strsql = strsql & ",'" & rsDsc("iden3l") & "','" & rsDsc("iden4") & "','" & rsDsc("iden4l") & "'"
                        strsql = strsql & ",'" & rsDsc("linkasoc") & "','0" & rsDsc("postxt") & "','0" & rsDsc("canttxt") & "','" & rsDsc("dattxt") & "')"

                        Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

                    End If
                    Call OFV.cerrarconsulta(rsDsc)

                  strsql = "SELECT objid From OBJLNK where iden2 = '" & sukey & "' and ltipo = 'KL' "
                  Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                  If Not rsDsc.EOF Then
                     wskllnk = rsDsc(0)
                  Else
                     wskllnk = 0
                  End If
                  Call OFV.cerrarconsulta(rsDsc)




                 end if




                 If clng(rs1(19)) > 0 and CCGIXLNK <> "N" Then
                    strsql = "SELECT * From OBJLNK where objid = " & rs1(19)
                    Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                    If Not rsDsc.EOF Then

                        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
                        strsql = strsql & ",iden3l,iden4,iden4l,linkasoc,postxt,canttxt,dattxt) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
                        strsql = strsql & "'" & nProced & "','" & sukey & "','" & rsDsc("iden3") & "',"
                        strsql = strsql & rsDsc("codobjetoL") & ",'" & rsDsc("tipobjetoL") & "','" & rsDsc("LTipo") & "','" & rsDsc("Lurl") & "','" & rsDsc("LurlDesc") & "'"
                        strsql = strsql & ",'" & rsDsc("iden3l") & "','" & rsDsc("iden4") & "','" & rsDsc("iden4l") & "'"
                        strsql = strsql & ",'" & rsDsc("linkasoc") & "','0" & rsDsc("postxt") & "','0" & rsDsc("canttxt") & "','" & rsDsc("dattxt") & "')"

                        Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

                    End If
                    Call OFV.cerrarconsulta(rsDsc)

                  strsql = "SELECT objid From OBJLNK where iden2 = '" & sukey & "' and ltipo <> 'KL'  "
                  Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                  If Not rsDsc.EOF Then
                     wsval3 = rsDsc(0)
                  Else
                     wsval3 = 0
                  End If
                  Call OFV.cerrarconsulta(rsDsc)

                  strsql = "Insert into Objaccion (Codobjeto,Tipobjeto,Codproced,Tipoaccion,modo,codcomp,"
                  strsql = strsql & "tipocomp,sectorcont,perfilcont,secuen,referencia,tiempo,copias,"
                  strsql = strsql & "referencia2,referencia3,codaccion,rellink,reflink,modi,relant,"
                  strsql = strsql & "CODSECUENPL,CRITICO,PRIORIDAD,descriacc,TIPPROCED"
                  strsql = strsql & ") Values("
                  strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "','" & nProced & "'," & rs1(4) & ",'" & rs1(5) & "',"
                  strsql = strsql & wskllnk & ",'" & rs1(7) & "'," & rs1(8) & ",'" & rs1(9) & "'," & rs1(10) & ",'"
                  strsql = strsql & rs1(11) & "'," & rs1(12) & "," & rs1(13)
                  strsql = strsql & ",'" & rs1(16) & "','" & rs1(17) & "','" & sukey & "'," & wsval3 & ","
                  strsql = strsql & "'" & rs1(20) & "',0,'" & rs1(0) & "',"
                  strsql = strsql & "'" & rs1(25) & "','" & rs1(26) & "','" & rs1(27) & "','" & rs1(28) & "','" & WKWTIPPROCED & "'"
                  strsql = strsql & ")"

                  Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              Else


                  strsql = "Insert into Objaccion (Codobjeto,Tipobjeto,Codproced,Tipoaccion,modo,codcomp,"
                  strsql = strsql & "tipocomp,sectorcont,perfilcont,secuen,referencia,tiempo,copias,"
                  strsql = strsql & "referencia2,referencia3,codaccion,rellink,reflink,modi,relant,"
                  strsql = strsql & "CODSECUENPL,CRITICO,PRIORIDAD,descriacc,TIPPROCED"
                  strsql = strsql & ") Values("
                  strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "','" & nProced & "'," & rs1(4) & ",'" & rs1(5) & "',"
                  strsql = strsql & wskllnk & ",'" & rs1(7) & "'," & rs1(8) & ",'" & rs1(9) & "'," & rs1(10) & ",'"
                  strsql = strsql & rs1(11) & "'," & rs1(12) & "," & rs1(13)
                  strsql = strsql & ",'" & rs1(16) & "','" & rs1(17) & "','" & sukey & "',0,'" & CCGIWRFLK & "',"
                  strsql = strsql & "0,'" & rs1(0) & "',"
                  strsql = strsql & "'" & rs1(25) & "','" & rs1(26) & "','" & rs1(27) & "','" & rs1(28) & "','" & WKWTIPPROCED & "'"
                  strsql = strsql & ")"

                  Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              End If



          '''''''''''''''

                     iaa = iaa + 1
                     End If
                     rs1.movenext
                  Loop
                  Call OFV.cerrarconsulta(rs1)
                  Call OFV.cerrarconn(CCGICN)

            '-----------------------------------------------
            '  Procedimientos - Acciones Alternativas
                  Set CCGICN = OFV.conectar(OFV.strconn2)
                  strsql = "SELECT OBJACCIONCONTROL.* FROM OBJACCION, OBJACCIONCONTROL where "
                  strsql = strsql & " OBJACCION.codaccion = OBJACCIONCONTROL.idaccion "
                  strsql = strsql & " and OBJACCION.codobjeto = " & CCGIOCOD
                  strsql = strsql & " and (OBJACCIONcontrol.modi = 0 or (OBJACCIONcontrol.modi <> 0"
                  strsql = strsql & " and OBJACCIONcontrol.tipmodi <> 'E')) order by OBJACCION.codaccion"

                  Set rsAlt = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

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
                        strsql = strsql & "continuador,modi,relant,"
                        strsql = strsql & "RELLINK,REFLINK) "
                        strsql = strsql & " Values ('" & nProced & "'," & rsAlt(2) & ",'" & sukey
                        strsql = strsql & "'," & rsAlt(4) & ",'" & rsAlt(5) & "','" & nconti & "',0,'" & rsAlt(0) & "',"
                        strsql = strsql & "'0','" & CCGIWRFLK & "')"
                        Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

                     End If
                     rsAlt.movenext

                  Loop
                  Call OFV.cerrarconsulta(rsAlt)
                  Call OFV.cerrarconn(CCGICN)






           case "BLKEST"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Editor Estructurado (CONTENIDO) ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Set CCGICN = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objestr Where codobjeto = " & CCGIOCOD & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
        EY = 0
        ei = 0
        Do Until rseditor.EOF
           EY = EY + 1
           ReDim Preserve aestr(ei)
           aestr(ei) = rseditor(0)
           ei = ei + 1
          strsql = "Insert Into Objestr (codobjeto,tipobjeto,contitulo,consecuen,modi,relant) Values ("
          strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "','" & rseditor(3) & "',"
          strsql = strsql & rseditor(4) & ",0,'" & rseditor(0) & "')"
          Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)

        strsql = "Select * From Objestr Where codobjeto = " & CCGINCOD & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
        ei = 0
        Do Until rseditor.EOF
           ReDim Preserve nestr(ei)
           nestr(ei) = rseditor(0)
           ei = ei + 1
           rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
     Call OFV.cerrarconn(CCGICN)

        Set CCGICN = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objestprty Where codobjeto = " & CCGIOCOD & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
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
          strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "'," & ndetcon & ",'" & rseditor(4) & "','"
          strsql = strsql & rseditor(5) & "','" & rseditor(6) & "','" & rseditor(7) & "','"
          strsql = strsql & rseditor(8) & "','" & rseditor(9) & "','" & rseditor(10) & "','"
          strsql = strsql & rseditor(11) & "','" & rseditor(12) & "','" & rseditor(13) & "','"
          strsql = strsql & rseditor(14) & "','" & rseditor(15) & "','" & rseditor(16) & "','"
          strsql = strsql & rseditor(17) & "','" & rseditor(18) & "','" & rseditor(19) & "','"
          strsql = strsql & rseditor(20) & "','" & rseditor(21) & "','" & rseditor(22)
          strsql = strsql & "')"
          Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
      Call OFV.cerrarconn(CCGICN)

        Set CCGICN = OFV.conectar(OFV.strconn2)

        strsql = " SELECT * FROM CSESTDET "
        strsql = strsql & " where codobjeto = " & CCGIOCOD
        strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
        strsql = strsql & " and not ( "
        strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
        strsql = strsql & " ) and stmodi = '' "
        strsql = strsql & " order by TSEC"

        Set rseditor = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
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

          detimg     = rseditor(9)
          detimgFlag = ""
          DETIMGID   = 0



        ' copia de link
          If rseditor(5) = "G" and IsNumeric(rseditor("RELLINK")) And rseditor("RELLINK") <> "0"  and CCGIXLNK <> "N" Then
             strsql = "SELECT * From OBJLNK where objid = 0" & rseditor("RELLINK")
             Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
            If Not rsDsc.EOF Then
              strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
              strsql = strsql & ",iden3l,iden4,iden4l,linkasoc,postxt,canttxt,dattxt) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
              strsql = strsql & "'" & rsDsc("iden1") & "','" & rsDsc("iden2") & "','" & rsDsc("iden3") & "',"
              strsql = strsql & rsDsc("codobjetoL") & ",'" & rsDsc("tipobjetoL") & "','" & rsDsc("LTipo") & "','" & rsDsc("Lurl") & "','" & rsDsc("LurlDesc") & "'"
              strsql = strsql & ",'" & rsDsc("iden3l") & "','" & rsDsc("iden4") & "','" & rsDsc("iden4l") & "'"
              strsql = strsql & ",'" & rsDsc("linkasoc") & "','0" & rsDsc("postxt") & "','0" & rsDsc("canttxt") & "','" & rsDsc("dattxt") & "')"

              Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              strsql = "SELECT * From OBJLNK where codobjeto=" & CCGINCOD & " and iden1='" & rsDsc("iden1") & "' and ltipo <> 'KL' "
              Set rsDsc2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
              if not rsDsc2.eof then
                 detimg     = rsDsc2(0)
                 detimgFlag = detimg
                 DETIMGID   = detimg

              else
                 detimg     = "9999"
                 detimgFlag = ""
                 DETIMGID   = 9999

              end if
              Call OFV.cerrarconsulta(rsDsc2)



            End If
            Call OFV.cerrarconsulta(rsDsc)
          End If

        ' copia de componente
          If rseditor(5) = "J" and IsNumeric(rseditor(9)) And rseditor(9) <> "0"  and CCGIXLNK <> "N" Then
             strsql = "SELECT * From OBJLNK where objid = 0" & rseditor(9)
             Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
            If Not rsDsc.EOF Then
              strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
              strsql = strsql & ",iden3l,iden4,iden4l,linkasoc,postxt,canttxt,dattxt) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
              strsql = strsql & "'" & rsDsc("iden1") & "','" & rsDsc("iden2") & "','" & rsDsc("iden3") & "',"
              strsql = strsql & rsDsc("codobjetoL") & ",'" & rsDsc("tipobjetoL") & "','" & rsDsc("LTipo") & "','" & rsDsc("Lurl") & "','" & rsDsc("LurlDesc") & "'"
              strsql = strsql & ",'" & rsDsc("iden3l") & "','" & rsDsc("iden4") & "','" & rsDsc("iden4l") & "'"
              strsql = strsql & ",'" & rsDsc("linkasoc") & "','0" & rsDsc("postxt") & "','0" & rsDsc("canttxt") & "','" & rsDsc("dattxt") & "')"

              Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              strsql = "SELECT * From OBJLNK where codobjeto=" & CCGINCOD & " and iden1='" & rsDsc("iden1") & "' and ltipo = 'KL' "
              Set rsDsc2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)


              detimg     = rsDsc2(0)
              detimgFlag = detimg
              DETIMGID   = 0
              Call OFV.cerrarconsulta(rsDsc2)
            End If
            Call OFV.cerrarconsulta(rsDsc)

          End If

            '
          strsql = "Insert Into Objestdet (codobjeto,tipobjeto,detcon,detsecuen,dettipo"
          strsql = strsql & ",detdepen,dettitulo,dettexto,detimagen,dettabla"
          strsql = strsql & ",detfont,detfsize,detfcolor,detbgcolor,detborde,detftipo"
          strsql = strsql & ",detsang,detindent1,detindent2,detvin,detvintipo,detvinindent"
          strsql = strsql & ",detnumer,detnumertipo,detnumerclase,detvinfont,detalign,dettextalign"
          strsql = strsql & ",detdefault,modi,RELLINK,REFLINK,RELANT"
          strsql = strsql & ") Values ("
          strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "'," & ndetcon & "," & rseditor(4)
          strsql = strsql & ",'" & rseditor(5) & "',"
          strsql = strsql & rseditor(6) & ",'" & rseditor(7) & "','" & replace(rseditor(8),"'","''") & "','"
          strsql = strsql & detimg & "','" & rseditor(10) & "','"
          strsql = strsql & rseditor(11) & "','" & rseditor(12) & "','" & rseditor(13) & "','"
          strsql = strsql & rseditor(14) & "','" & rseditor(15) & "','" & rseditor(16) & "','"
          strsql = strsql & rseditor(17) & "','" & rseditor(18) & "','" & rseditor(19) & "','"
          strsql = strsql & rseditor(20) & "','" & rseditor(21) & "','" & rseditor(22) & "','"
          strsql = strsql & rseditor(23) & "','" & rseditor(24) & "','" & rseditor(25) & "','"
          strsql = strsql & rseditor(26) & "','" & rseditor(27) & "','" & rseditor(28) & "','"
          strsql = strsql & rseditor(31)
          strsql = strsql & "',0,'" & DETIMGID & "','" & rseditor(33) & "','" & rseditor(0) & "')"
          'aca contenido

          Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)


          ' cambio el iden1 y iden2
          If detimgFlag <> "" Then
              strsql = "select objid from objestdet where (RELLINK=" & detimgFlag & " or detimagen = '" & detimgFlag & "' ) and codobjeto=" & CCGINCOD
              Set rsDsc2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
              if not rsDsc2.eof then
                 strsql = "update objlnk set iden1=" & rsDsc2(0) & ",iden2=" & rsDsc2(0) & " where objid=" & detimgFlag & " and codobjeto=" & CCGINCOD
                 Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
              end if
              Call OFV.cerrarconsulta(rsDsc2)
          End If

          rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
      Call OFV.cerrarconn(CCGICN)

        Set CCGICN = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objestdet Where codobjeto = " & CCGINCOD & " order by objid"
        Set rseditor = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

        EY = 0
        ei = 0
        Do Until rseditor.EOF
           EY = EY + 1
           ei = ei + 1

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

          Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          rseditor.movenext
        Loop
      Call OFV.cerrarconsulta(rseditor)
      Call OFV.cerrarconn(CCGICN)


           case "BLKCNT"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' CONTROLES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         CY = 0
        Set CCGICN = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objcontrol Where codobjeto = " & CCGIOCOD & " order by objid"
        Set rsControl = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
        Do Until rsControl.EOF
          strsql = "Insert Into Objcontrol (codobjeto,tipobjeto,codsector,frecuencia,codperfil,referencia,modi,relant) Values ("
          strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "'," & rsControl(3) & ",'"
          strsql = strsql & rsControl(4) & "','" & rsControl(5) & "','" & rsControl(6) & "',0,'" & rsControl(0) & "')"
          Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          ReDim Preserve Ocontrol(CY)
          Ocontrol(CY) = rsControl(0)
          CY = CY + 1
          rsControl.movenext
        Loop
        Call OFV.cerrarconsulta(rsControl)
        CY = 0
        strsql = "Select objid From Objcontrol Where codobjeto = " & CCGINCOD & " order by objid"
        Set rsControlb = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
        Do Until rsControlb.EOF
          ReDim Preserve ycontrol(CY)
          ycontrol(CY) = rsControlb(0)
          CY = CY + 1
          rsControlb.movenext
        Loop
        Call OFV.cerrarconsulta(rsControlb)

          strsql = "Select * From Objcontrol2 Where codobjeto = " & CCGIOCOD
          strsql = strsql & " and (Objcontrol2.modi = 0 or (Objcontrol2.modi <> 0"
          strsql = strsql & " and Objcontrol2.tipmodi <> 'E')) order by control"
          Set rsControl2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          Do Until rsControl2.EOF
             nControl = 0
             For CY = 0 To UBound(ycontrol)
                 If CStr(rsControl2("control")) = CStr(Ocontrol(CY)) Then
                    nControl = ycontrol(CY)
                 End If
             Next

        ' copia de link
          If IsNumeric(rsControl2("RELLINK")) And rsControl2("RELLINK") <> "0"  and CCGIXLNK <> "N" Then
             strsql = "SELECT * From OBJLNK where objid = '" & rsControl2("RELLINK") & "'"
             Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
            If Not rsDsc.EOF Then

              strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
              strsql = strsql & ",iden3l,iden4,iden4l,linkasoc,postxt,canttxt,dattxt) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
              strsql = strsql & "'" & rsDsc("iden1") & "','" & rsDsc("iden2") & "','" & rsDsc("iden3") & "',"
              strsql = strsql & rsDsc("codobjetoL") & ",'" & rsDsc("tipobjetoL") & "','" & rsDsc("LTipo") & "','" & rsDsc("Lurl") & "','" & rsDsc("LurlDesc") & "'"
              strsql = strsql & ",'" & rsDsc("iden3l") & "','" & rsDsc("iden4") & "','" & rsDsc("iden4l") & "'"
              strsql = strsql & ",'" & rsDsc("linkasoc") & "','0" & rsDsc("postxt") & "','0" & rsDsc("canttxt") & "','" & rsDsc("dattxt") & "')"

              Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              strsql = "SELECT * From OBJLNK where codobjeto=" & CCGINCOD & " and iden1='" & rsDsc("iden1") & "'"
              Set rsDsc2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              detimg     = rsDsc2(0)
              detimgFlag = detimg
              Call OFV.cerrarconsulta(rsDsc2)
            End If
            Call OFV.cerrarconsulta(rsDsc)
          Else
                detimg     = 0
                detimgFlag = ""
          End If

            strsql = "Insert Into Objcontrol2 (codobjeto,tipobjeto,control,descri,secuencia,RELLINK,REFLINK,modi,relant) Values ("
            strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "'," & nControl & ",'" & rsControl2(4)
            strsql = strsql & "'," & rsControl2(5) & ",'" & detimg & "','" & rsControl2(9) & "',0,'" & rsControl2(0) & "')"
            Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

          ' cambio el iden1 y iden2
          If detimgFlag <> "" Then
              strsql = "select objid from Objcontrol2 where RELLINK=" & detimgFlag & " and codobjeto=" & CCGINCOD
              Set rsDsc2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
              strsql = "update objlnk set iden1=" & nControl & ",iden2=" & rsDsc2(0) & " where objid=" & detimgFlag & " and codobjeto=" & CCGINCOD

              Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
              Call OFV.cerrarconsulta(rsDsc)
              Call OFV.cerrarconsulta(rsDsc2)
          End If


            rsControl2.movenext
          Loop
          Call OFV.cerrarconsulta(rsControl2)


      Call OFV.cerrarconn(CCGICN)




           case "BLKESQ"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' ESQUEMAS CONTABLES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         CZ = 0
         Set CCGICN = OFV.conectar(OFV.strconn2)
        strsql = "Select * From Objesquema Where codobjeto = " & CCGIOCOD & " order by objid"
        Set rCCGITESQ = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
        i = 0
        Do Until rCCGITESQ.EOF
          Neven = 0
          For ei = 0 To UBound(aevento)
              If CStr(aevento(ei)) = CStr(rCCGITESQ(3)) Then
                 Neven = nevento(ei)
                 Exit For
              End If
          Next
          strsql = "Insert Into Objesquema (codobjeto,tipobjeto,codevento,sectororigen,"
          strsql = strsql & "sectordestino,refesq,modi,relant) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
          strsql = strsql & Neven & "," & rCCGITESQ(4) & "," & rCCGITESQ(5) & ",'" & rCCGITESQ(6) & "',0,'" & rCCGITESQ(0) & "')"
          Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          ReDim Preserve Oesque(CZ)
          Oesque(CZ) = rCCGITESQ(0)
          CZ = CZ + 1
          rCCGITESQ.movenext
          i = i + 1
          Loop
          Call OFV.cerrarconsulta(rCCGITESQ)
          CZ = 0
          strsql = "Select objid From Objesquema Where codobjeto = " & CCGINCOD & " order by objid"
          Set rCCGITESQb = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          Do Until rCCGITESQb.EOF
          ReDim Preserve Nesque(CZ)
          Nesque(CZ) = rCCGITESQb(0)
          CZ = CZ + 1
          rCCGITESQb.movenext
          Loop
          Call OFV.cerrarconsulta(rCCGITESQb)
          strsql = "Select * From Objesquema2 Where codobjeto = " & CCGIOCOD
          strsql = strsql & " and (Objesquema2.modi = 0 or (Objesquema2.modi <> 0"
          strsql = strsql & " and Objesquema2.tipmodi <> 'E')) order by codesquema "
          Set rCCGITESQ2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
          Do Until rCCGITESQ2.EOF
             CCGICESQ = 0
             For CZ = 0 To UBound(Nesque)
                 If CStr(rCCGITESQ2("codesquema")) = CStr(Oesque(CZ)) Then
                    CCGICESQ = Nesque(CZ)
                 End If
             Next

        ' copia de link
          If IsNumeric(rCCGITESQ2("RELLINK")) And rCCGITESQ2("RELLINK") <> "0"  and CCGIXLNK <> "N" Then
             strsql = "SELECT * From OBJLNK where objid = '" & rCCGITESQ2("RELLINK") & "'"
             Set rsDsc = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
            If Not rsDsc.EOF Then

              strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
              strsql = strsql & ",iden3l,iden4,iden4l,linkasoc,postxt,canttxt,dattxt) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
              strsql = strsql & "'" & rsDsc("iden1") & "','" & rsDsc("iden2") & "','" & rsDsc("iden3") & "',"
              strsql = strsql & rsDsc("codobjetoL") & ",'" & rsDsc("tipobjetoL") & "','" & rsDsc("LTipo") & "','" & rsDsc("Lurl") & "','" & rsDsc("LurlDesc") & "'"
              strsql = strsql & ",'" & rsDsc("iden3l") & "','" & rsDsc("iden4") & "','" & rsDsc("iden4l") & "'"
              strsql = strsql & ",'" & rsDsc("linkasoc") & "','0" & rsDsc("postxt") & "','0" & rsDsc("canttxt") & "','" & rsDsc("dattxt") & "')"

              Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              strsql = "SELECT * From OBJLNK where codobjeto=" & CCGINCOD & " and iden1='" & rsDsc("iden1") & "'"
              Set rsDsc2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

              detimg     = rsDsc2(0)
              detimgFlag = detimg
              Call OFV.cerrarconsulta(rsDsc2)
            End If
            Call OFV.cerrarconsulta(rsDsc)
          Else
                detimg     = 0
                detimgFlag = ""
          End If


            strsql = "Insert Into Objesquema2 (codobjeto,tipobjeto,codesquema,codcuenta,sentido,refesq2,RELLINK,REFLINK,modi,relant) "
            strsql = strsql & "Values (" & CCGINCOD & ",'" & CCGINTIP & "'," & CCGICESQ & ","
            strsql = strsql & rCCGITESQ2(4) & ",'" & rCCGITESQ2(5) & "','" & rCCGITESQ2(6) & "','" & detimg & "','" & rCCGITESQ2(10) & "',0,'" & rCCGITESQ2(0) & "')"
            Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)

          ' cambio el iden1 y iden2
          If detimgFlag <> "" Then
              strsql = "select objid from Objesquema2 where RELLINK=" & detimgFlag & " and codobjeto=" & CCGINCOD
              Set rsDsc2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
              if not rsdsc2.eof then
                 strsql = "update objlnk set iden1=" & CCGICESQ & ",iden2=" & rsDsc2(0) & " where objid=" & detimgFlag & " and codobjeto=" & CCGINCOD
                 call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
              end if
              Call OFV.cerrarconsulta(rsDsc2)
          End If

            rCCGITESQ2.movenext
          Loop
          Call OFV.cerrarconsulta(rCCGITESQ2)


      Call OFV.cerrarconn(CCGICN)




           case "BLKDOC"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DOCUMENTOS ASOCIADOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                   Set CCGICN = OFV.conectar(OFV.strconn2)
                   strsql = "Select * From Objhtm Where codobjeto = " & CCGIOCOD & " order by objid"
                   Set rshtm = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                   EY = 0
                   ei = 0
                   Do Until rshtm.EOF
                      EY = EY + 1
                      ei = ei + 1
                      strsql = "Insert Into Objhtm (codobjeto,tipobjeto,secuen,descri,archivo,TIPO,directorio,modi,relant) Values ("
                      strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "'," & rshtm(3) & ",'"
                      strsql = strsql & rshtm(4) & "','" & rshtm(6) & "','" & rshtm(7) & "','" & rshtm(8) & "',0,'" & rshtm(0) & "')"
                      Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                      rshtm.movenext
                   Loop
                   Call OFV.cerrarconsulta(rshtm)
                   Call OFV.cerrarconn(CCGICN)


           case "BLKTAB"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' TABLAS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                   Set CCGICN = OFV.conectar(OFV.strconn2)
                   strsql = "Select * From Objtabdef Where codobjeto = " & CCGIOCOD & " order by objid"
                   Set rstablas = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                   EY = 0
                   ei = 0
                   Do Until rstablas.EOF
                      EY = EY + 1
                      ReDim Preserve aTabla(ei)
                      aTabla(ei) = rstablas(0)
                      ei = ei + 1
                      strsql = "Insert Into Objtabdef (codobjeto,tipobjeto,tabperfil,tabtitulo,tabsecuen,MODI,RELANT) Values ("
                      strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "','" & rstablas(3) & "','"
                      strsql = strsql & rstablas(4) & "'," & rstablas(5) & ",0,'" & rstablas(0) & "')"
                      Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                      rstablas.movenext
                   Loop
                   Call OFV.cerrarconsulta(rstablas)

                   strsql = "Select * From Objtabdef Where codobjeto = " & CCGINCOD & " order by objid"
                   Set rstablas = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                   ei = 0
                   Do Until rstablas.EOF
                      ReDim Preserve nTabla(ei)
                      nTabla(ei) = rstablas(0)
                      ei = ei + 1
                      rstablas.movenext
                   Loop
                   Call OFV.cerrarconsulta(rstablas)
                   Call OFV.cerrarconn(CCGICN)

                   Response.write "<p>  Tablas: " & EY & " </p>"

                   Set CCGICN = OFV.conectar(OFV.strconn2)
                   strsql = "Select * From Objtabblk Where codobjeto = " & CCGIOCOD & " order by objid"
                   Set rstablas = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
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

                      strsql = "Insert Into Objtabblk (codobjeto,tipobjeto,tabdef,tabblksec,tabblkdescri,MODI,RELANT) Values ("
                      strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "'," & Ndef & ","
                      strsql = strsql & rstablas(4) & ",'" & rstablas(5) & "',0,'" & rstablas(0) & "')"
                      Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                     'response.write strsql
                      rstablas.movenext
                   Loop
                   Call OFV.cerrarconsulta(rstablas)

                   strsql = "Select * From Objtabblk Where codobjeto = " & CCGINCOD & " order by objid"
                   Set rstablas = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                   ei = 0
                   Do Until rstablas.EOF
                      ReDim Preserve nTablab(ei)
                      nTablab(ei) = rstablas(0)
                      ei = ei + 1
                      rstablas.movenext
                   Loop
                   Call OFV.cerrarconsulta(rstablas)
                   Call OFV.cerrarconn(CCGICN)

                   Response.write "<p>  Bloques: " & EY & " </p>"


                   Set CCGICN = OFV.conectar(OFV.strconn2)
                   strsql = "Select * From Objtabcon Where codobjeto = " & CCGIOCOD & " order by objid"
                   Set rstablas = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
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
                     strsql = strsql & ",TABCONID,MODI,RELANT) Values ("
                     strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "'," & Nblk & ","
                     strsql = strsql & rstablas(4) & ",'" & rstablas(5) & "','"
                     strsql = strsql & rstablas(6) & "','" & rstablas(7) & "','" & rstablas(8) & "','"
                     strsql = strsql & rstablas(9) & "','" & rstablas(10) & "','" & rstablas(11) & "','"
                     strsql = strsql & rstablas(12) & "','" & rstablas(13) & "','" & rstablas(14) & "','"
                     strsql = strsql & rstablas(15) & "','" & rstablas(16) & "','" & rstablas(17) & "','"
                     strsql = strsql & rstablas(18) & "','" & rstablas(19) & "','" & rstablas(20) & "','"
                     strsql = strsql & rstablas(21) & "','" & rstablas(22) & "','" & rstablas(23) & "','"
                     strsql = strsql & rstablas(24) & "','" & rstablas(25)
                     strsql = strsql & "','TB" & GTkey(CCGINCOD) & EY & "',0,'" & rstablas(26) & "')"
                     Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                     rstablas.movenext
                   Loop
                   Call OFV.cerrarconsulta(rstablas)
                   Call OFV.cerrarconn(CCGICN)

                   Set CCGICN = OFV.conectar(OFV.strconn2)
                   strsql = "Select * From ObjtabCMB Where codobjeto = " & CCGIOCOD & " order by objid"
                   Set rstablas = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                   EY = 0

                   Do Until rstablas.EOF
                      EY = EY + 1

                      NTCONID = "0000"

                      strsql = "Select TABCONID From Objtabcon Where codobjeto = " & CCGINCOD & " AND RELANT = '" & rstablas(4) & "' "
                      Set rstablas2 = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                      IF NOT rstablas2.EOF THEN   NTCONID = rstablas2(0)
                      Call OFV.cerrarconsulta(rstablas2)


                      strsql = "Insert Into ObjtabCMB (codobjeto,tipobjeto,tabblk,tabOBJ,tabtipo"
                      strsql = strsql & ",tabCOL,CMBCOL,CMBROW,CMBSTYL,CMBFONT,CMBSIZE"
                      strsql = strsql & ",CMBCOLOR,CMBFONDO,CMBBORDE) Values ("
                      strsql = strsql & CCGINCOD & ",'" & CCGINTIP & "',0,'"
                      strsql = strsql & NTCONID & "','" & rstablas(5) & "','"
                      strsql = strsql & rstablas(6) & "','" & rstablas(7) & "','" & rstablas(8) & "','"
                      strsql = strsql & rstablas(9) & "','" & rstablas(10) & "','" & rstablas(11) & "','"
                      strsql = strsql & rstablas(12) & "','" & rstablas(13) & "','" & rstablas(14) & "')"
                      Call OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
                      rstablas.movenext
                   Loop
                   Call OFV.cerrarconsulta(rstablas)
                   Call OFV.cerrarconn(CCGICN)




           case "BLKFAQ"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' REDIRECCION DE DOCUMENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      Set CCGICN = OFV.conectar(OFV.strconn2)
      strsql = "Select * From Objfaq Where codobjeto = " & CCGIOCOD
      Set rshtm = OFV.crearconsultaEx(strsql, CCGICN, 1, volver)
      if not rshtm.eof then
         EY = 0

         strsql = "Insert into Objfaq (codobjeto,tipobjeto"
         For ii = 3 To (rshtm.fields.Count-1)
             strsql = strsql & "," & rshtm(ii).Name
         Next
         strsql = strsql & ",modi,relant) values (" & CCGINCOD & ",'" & CCGINTIP & "'"


         Do Until rshtm.EOF
            strsql2 = ""
            EY = EY + 1
            For ii = 3 To (rshtm.fields.Count-1)
               sInput = rshtm(ii)
               If rshtm(ii).Type = 131 Then sInput = Replace(rshtm(ii), ",", ".")
               strsql2 = strsql2 & ",'" & sInput & "'"
            Next

            Call OFV.crearconsultaEx(strsql & strsql2 & ",0,'" & rshtm(0) & "') ", CCGICN, 1, volver)
            rshtm.movenext
         Loop
         Call OFV.cerrarconsulta(rshtm)
         Call OFV.cerrarconn(CCGICN)

      end if



           case "BLKCON"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' CONSIDERACIONES GENERALES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Set CCGICN = OFV.conectar(OFV.strconn2)

        Strsql = "Select * From OBJNWEDT Where codobjeto=" & CCGIOCOD & " Order By consecuen"
        set rsProced = ofv.crearconsultaEx(StrSql,CCGICN,1,parametros)
        if  rsProced.eof then
            xxxx = NWEDTVER(1,CCGINCOD,CCGIOCOD,CCGINTIP)
        else
            xxxx = NWEDTVER(2,CCGINCOD,CCGIOCOD,CCGINTIP)
        end if

        Call OFV.cerrarconn(CCGICN)




    end select



'''''''''''



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


function NWEDTVER(xcpopc,xpcodigo,xvcodigo,xptipo)

  XNVER = "N"
  XNSEC = 0

  select case xcpopc

  case 1

  Xfer        = server.mappath(request.servervariables("SCRIPT_NAME"))
  XferTarget0 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetasCMD\"
  XferTarget1 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "carpetaObjetivos\"
  XferTarget2 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaAlcance\"
  XferTarget3 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaConsideracion\"

  abrir1      = XferTarget1 & RtfObjetivo
  abrir2      = XferTarget2 & RtfAlcance
  abrir3      = XferTarget3 & RtfConsid

  set fff     = server.createobject("scripting.filesystemobject")


        Set xcn = OFV.conectar(OFV.strconn2)

'''''''''''''' Objetivo
''''
  Strsql  = "Select ARCHIVO,titulo From ObjMain where CodObjeto=" & xvcodigo
  set rs1 = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
  if  not rs1.eof then

      XNVER       = "S"
      XNSEC       = XNSEC + 10
      RtfObjetivo = rs1(0)
      TITObjetivo = rs1(1)

      XNNWCOD     = GETNWEN(xpcodigo,xptipo) & XNSEC
      XNABRIR     = XferTarget1 & RtfObjetivo
      xarchivo    = ofv.convertircar(XNABRIR,chr(34)," ")
      xxx         = ""

    if fff.FileExists(trim(Xarchivo)) then
       Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
       do  while fff1.AtEndOfstream <> true
           xxx = xxx & fff1.readline
       loop
       fff1.close

       set XNPSR  = server.createobject("CorRtfParser.RtfParser")
       if xxx <> "" then
          XNPSRFL    =  XNPSR.RtfToHtml(xxx)
       end if
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql     = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,coCCGINCOD,modi,relant) values(" & xpcodigo & ",'"
       Strsql     = Strsql & xptipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "',0,'0')"
       call ofv.crearconsultaEx(StrSql,xcn,1,parametros)

    end if
  end if
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Alcance
''''
  Strsql  = "Select ARCHIVO,titulo From ObjAlcance where CodObjeto=" & xvcodigo
  set rs1 = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
  if  not rs1.eof then

      XNVER       = "S"
      XNSEC       = XNSEC + 10
      RtfObjetivo = rs1(0)
      TITObjetivo = rs1(1)

      XNNWCOD     = GETNWEN(xpcodigo,xptipo) & XNSEC
      XNABRIR     = XferTarget2 & RtfObjetivo
      xarchivo    = ofv.convertircar(XNABRIR,chr(34)," ")
      xxx         = ""

    if fff.FileExists(trim(Xarchivo)) then
       Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
       do  while fff1.AtEndOfstream <> true
           xxx = xxx & fff1.readline
       loop
       fff1.close

       set XNPSR  = server.createobject("CorRtfParser.RtfParser")

       if xxx<> "" then
         XNPSRFL    =  XNPSR.RtfToHtml(xxx)
       end if
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql     = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,coCCGINCOD,modi,relant) values(" & xpcodigo & ",'"
       Strsql     = Strsql & xptipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "',0,'0')"
       call ofv.crearconsultaEx(StrSql,xcn,1,parametros)

    end if
  end if
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Consideraciones generales
''''
  Strsql  = "Select ARCHIVO,titulo From ObjConsid where CodObjeto=" & xvcodigo
  set rs1 = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
  if  not rs1.eof then

      XNVER       = "S"
      XNSEC       = XNSEC + 10
      RtfObjetivo = rs1(0)
      TITObjetivo = rs1(1)

      XNNWCOD     = GETNWEN(xpcodigo,xptipo) & XNSEC
      XNABRIR     = XferTarget3 & RtfObjetivo
      xarchivo    = ofv.convertircar(XNABRIR,chr(34)," ")
      xxx         = ""

    if fff.FileExists(trim(Xarchivo)) then
       Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
       do  while fff1.AtEndOfstream <> true
           xxx = xxx & fff1.readline
       loop
       fff1.close

       set XNPSR  = server.createobject("CorRtfParser.RtfParser")

       if xxx <> "" then
         XNPSRFL    =  XNPSR.RtfToHtml(xxx)
       end if
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql     = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,CONCODIGO,modi,relant) values(" & xpcodigo & ",'"
       Strsql     = Strsql & xptipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "',0,'0')"
       call ofv.crearconsultaEx(StrSql,xcn,1,parametros)

    end if
  end if
  call ofv.cerrarconsulta(rs1)

  Call OFV.cerrarconn(xcn)

 case 2

        Xfer         = server.mappath(request.servervariables("SCRIPT_NAME"))
        Set xcn      = OFV.conectar(OFV.strconn2)

        Strsql       = "Select * From OBJNWEDT Where codobjeto=" & xvcodigo & " Order By consecuen"
        set rsProced = ofv.crearconsultaEx(StrSql,xcn,1,parametros)

        do until rsProced.eof

        XNSEC        = rsProced(4)
        RtfObjetivo  = rsProced(5)
        TITObjetivo  = rsProced(3)

        XNNWCOD      = GETNWEN(xpcodigo,xptipo) & XNSEC

        strsql       = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,coNCODIGO,modi,relant) values(" & xpcodigo & ",'"
        Strsql       = Strsql & xptipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "',0,'" & rsProced(0) & "')"

        call ofv.crearconsultaEx(StrSql,xcn,1,parametros)

        On Error Resume Next

        sOrigen      =RtfObjetivo & ".CSE"
        sDestino     =  XNNWCOD     & ".CSE"

        if SESSION("FILETIPO")="L" THEN
        Set fso = Server.CreateObject("scripting.filesystemobject")
         sPath = Server.mappath(Request.servervariables("SCRIPT_NAME"))
		 sPath = Mid(sPath, 1, InStrRev(sPath, "\", -1))
         sPatH = sPath & "CarpetaSCMD\"

        fso.CopyFile sPatH & sOrigen, sPatH & sDestino
        ELSE

        url=session ("DirUpload") &"/asp/copiar.asp?file1=" & sOrigen & "&file2=" & sDestino

        call callurl("",url)
        END IF
         rsProced.movenext
         Icons=Icons+1
     loop
     Response.write "<p>  Consideraciones Generales: " & Icons &  " </p>"
     Response.write "tiempo inicial: " & ntime0 & " final: " & Minute(Now) & ":" & Second(Now)
     Call OFV.cerrarconsulta(rsProced)

      Call OFV.cerrarconn(xcn)

 end select


end function



%>