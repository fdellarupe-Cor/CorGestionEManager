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
  MOVERESQUEMA= Request.querystring("MOVERESQUEMA")
  nCodigo    = Request.querystring("Codigo")
  sTipo      = Request.querystring("Tipo")
  sPerfil    = Request.querystring("Perfil")
  bCopiar    = Request.querystring("Copiar")
  sOpcion    = Request.querystring("Opcion")
  sEmpresa   = Request.querystring("empresa")
  sUsuario   = Session("Usuario")
  sTipobjeto = sOpcion
  volver = ""
  ESQSUP=Request.querystring("ESQSUP")

  IF sOpcion="TODOS" THEN ESQSUP=sOpcion
'''''''''''''
  'RESPONSE.WRITE REQUEST.QUERYSTRING

  if moveresquema="SI" THEN
    sREADONLY= " readonly "
  ELSE
    sREADONLY= ""
  END IF

   If bCopiar = "SI" Then


        Server.scripttimeout = 600
        Response.write OFV.FormBody

        Response.write "<font face='verdana' size=3 color=#336699>"

        Response.write "<center><p style='border-style=solid;border-width:1;border-color:#eeeeee;' ><b>  Registro de la Copia </b></p><br>"

        styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; "
        styl = styl & "SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa;"
        styl = styl & "SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"


        Response.write "<div id='mt2'  style='border-style=solid;border-width:2;border-color:#eeeeee;height:350px;width:400px;overflow:auto;" & styl & "' >"

        Response.write "<font face='verdana' size=2 color=#336699>"

        ntime0     = Minute(Now) & ":" & Second(Now)
        Response.write "tiempo proc inicio: " & ntime0
        xCodigo    = nCodigo
        xtipobjeto = sTipobjeto

        sCodiden   = UCase(Request.Form("codiden"))
        sDescri    = UCase(Request.Form("descri"))
        sEsquema   = Request.Form("Esquema")

        'response.write sEsquema & "2222"
        i          = InStr(sEsquema, "_")
        nEsquema   = Left(sEsquema, i - 1)
        sEsquema   = Mid(sEsquema, i + 1)



  CCGIOCOD = nCodigo
  CCGIOTIP = sTipo

  CCGINCOD = 0
  CCGINTIP = sTipobjeto

  CCGIEMPR = sEmpresa
  CCGINPRF = sPerfil
  CCGIDESC = sDescri
  CCGICDID = sCodiden
  CCGICESQ = nEsquema
  CCGITESQ = sEsquema


  CCGESQSUP= ESQSUP




    Set cn = OFV.conectar(OFV.strconn2)

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

       strsql = "Select * From Objperfildoc Where codobjeto =" & xCodigo
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
          strsql = strsql & "and Tipobjeto.tipobjeto = '" & xtipobjeto & "'"
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


SELECT CASE MOVERESQUEMA

CASE "SI"

   xxxx = CopyGenFn("","MOVE")

CASE   ELSE ' COPIAR


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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
END SELECT

        Response.write "tiempo  fin: " & Minute(Now) & ":" & Second(Now) & "<br>"



        Response.write "</div>"

        Response.write "<center>"

        sInput = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "")
        Response.write OFV.GenForm("Copiar", "", "vuelve.asp", "", "post", "", sInput, "si")
        Response.write "</center>"

      sHTML = "</body></html>"


    Else
      sHTML = OFV.FormBody & corstyler2()
   'Objetos

  Set cn = OFV.conectar(OFV.strconn2)
  SELECT CASE ESQSUP



  case "Y" ' SUBCARPETA INCLUIDA
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





	'''sEmpresa

	        if  sEmpresa = "" then sEmpresa = 1
	        bempresa = sEmpresa
	        sLink    = "&nbsp;"

	        Fila = ""

	  IF  Session("EMPMUN") = "S" THEN

	        FOR IAEMP=1 TO TAEMP

	            IF ATREMP(IAEMP,2) = 1 THEN
	               IF ATREMP(IAEMP,0) = sEmpresa THEN
	                  bempresa = ATREMP(IAEMP,4)
	                  sLink    = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
	                  exit for
	               END IF
	            END IF

	        NEXT

	       'empresa
	        Celda = OFV.GenCelda("", "ttt ", "", "", "", "", "", sLink, "si")
	        sInput = "&nbsp;&nbsp;" & bempresa
	        Celda = Celda & OFV.GenCelda("", "ttt", "", "", "", "", "", sInput, "si")
	        Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

	        Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;&nbsp;", "si")
	        Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

	   end if



          Strsql = "Select * From Subtema WHERE tipobjeto='"& stipo &"' and (aplic = '" & wkaplic & "' or aplic = 'GLB') Order By descri"
		    ' response.write strsql
   	      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		  if not rs.eof then sNombredOC = rs("descri") &":"
		  call ofv.cerrarconsulta(rs)

		 strsql = "SELECT DESCRI FROM PRODSERV where codprodserv = " & nCodigo
		  Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		  if not rs.eof then sNombredOC = sNombredOC & rs(0)
		  call ofv.cerrarconsulta(rs)




	      'Identificacion
	      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "NOMBRE", "si")
	      sInput = sNombredOC
	      Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
	      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




	      'Esquema superior
	      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Mover a Tema: ", "si")
	      sInput = "<font class=at><select name=Esquema>"
	      If aTipos <> ")" Then
	         sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
	         sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
	         sInput = sInput & "<--          Seleccione un Tema          -->"
	         sInput = sInput & "</OPTION>"

             strsql="select * from tema order by  descri"
             Set rsTipo2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
	         Do Until rsTipo2.EOF

	            sInput = sInput & "<option value='" & rsTipo2(0) & "_" & rsTipo2(0) &"' "
	            sInput = sInput & " >" & rsTipo2("descri") & "</option>"


	            rsTipo2.movenext
	        Loop
	      End If


	      sInput = sInput & "</select></font>"
	      Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
	      Fila   = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

	      Tabla  = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")


	      Tabla      = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")
	      Fila       = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

	      sInput     = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "this.style.display='none';")
	      Celda      = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
	      sAccion    = "Copiar2.asp?Codigo=" & nCodigo & "&Tipo=" & sTipo & "&Perfil=" & sPerfil & "&empresa=" & sEmpresa
	      sAccion    = sAccion & "&Opcion=" & sOpcion & "&Copiar=SI" & "&moveresquema=" & moveresquema & "&ESQSUP=" & ESQSUP
	      Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

	      sFunc      = "window.close()"
	      sInput     = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
	      Celda      = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
	      Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
	      Fila       = Formulario & OFV.GenFRow()

	      Tabla      = Fila & OFV.GenFTabla()

	      sHTML      = sHTML & "<br><br><br><br><center>" & Tabla & "</center>"
	      sHTML      = sHTML & "</body></html>"


 CASE "TODOS" ' TODOS MENOS LA SUBCARPETA

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
		       IF  rsObj.EOF THEN
		        XTipo=sTipo
		       ELSE
		       XTipo  = rsObj(2)
		       END IF
		       strsql = "Select codobjmaster,tipobjmaster From relaobj Where codojasoc =" & nCodigo
		       strsql = strsql & " and tipobjasoc = '" & XTipo	 & "' "
		      ' response.write strsql

		       Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)



		       bSector = "S/D"
		       bControl = "N"
		       Nsect = 0
		       If Not rsNew.EOF Then
		       Nsect = rsNew(0)
		       For inp = 0 To UBound(asector)
		           If asector(inp) = UCase(rsNew(1)) Then
		              bSector  = nsector(inp)
		              bControl = "S"
		              Exit For
		            End If
		       Next

		       If bControl = "N" Then
		          For inp = 0 To UBound(aevento)
		           If aevento(inp) = UCase(rsNew(1)) Then
		              bSector  = nevento(inp)
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
		          strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
		          strsql = strsql & "Tema.codtema = Prodserv.codtema and prodserv.codprodserv = " & rsNew(0)
		          Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
		          If Not rsNew.EOF Then
		             bSector = bSector & ": " & rsNew(0) & "/" & rsNew(2)
		          End If

		       End If

		       Call OFV.cerrarconsulta(rsNew)
		       End If

		       bEvento = "S/D"




		'''sEmpresa

		        if  sEmpresa = "" then sEmpresa = 1
		        bempresa = sEmpresa
		        sLink    = "&nbsp;"

		        Fila = ""

		  IF  Session("EMPMUN") = "S" THEN

		        FOR IAEMP=1 TO TAEMP

		            IF ATREMP(IAEMP,2) = 1 THEN
		               IF ATREMP(IAEMP,0) = sEmpresa THEN
		                  bempresa = ATREMP(IAEMP,4)
		                  sLink    = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		                  exit for
		               END IF
		            END IF

		        NEXT

		       'empresa
		        Celda = OFV.GenCelda("", "ttt ", "", "", "", "", "", sLink, "si")
		        sInput = "&nbsp;&nbsp;" & bempresa
		        Celda = Celda & OFV.GenCelda("", "ttt", "", "", "", "", "", sInput, "si")
		        Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

		        Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;&nbsp;", "si")
		        Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

		   end if

		       'origen
		      '  Celda = "" 'OFV.GenCelda("", "tt", "", "", "", "", "", " Copiar:" , "si")
		      '  sInput = "" 'bEvento & ": " & rsObj(1) & " - " & rsObj(3)
		      '  For inp = 0 To UBound(aestr)
		      '      If aestr(inp) = UCase(rsObj(9)) Then
		      '         sInput = sInput & " (" & nestr(inp) & ")"
		       '        Exit For
		       '     End If
		       ' Next
		        'Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
		        'Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




		       'Superior
		       ' Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Perteneciente a:", "si")
		        'sInput = bSector
		       ' Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
		       ' Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

		        'Como
		        Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Como:", "si")
		        sInput = sOpcion
		        For inp = 0 To UBound(aevento)
		            If aevento(inp) = UCase(sOpcion) Then
		               sInput = nevento(inp)
		               Exit For
		            End If
		        Next
		     '   Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
		     '   Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

		      'Descripcion
		      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Descripcion", "si")
		      sInput = OFV.GenerarInput(rsObj(3).Name & sREADONLY, rsObj(3), "agregar", "Copiar", "80", "dt", "")
		      Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
		      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

		      'Identificacion
		      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Identificacion", "si")
		      sInput = OFV.GenerarInput(rsObj(1).Name & sREADONLY, rsObj(1), "agregar", "copiar", "20", "dt", "")
		      Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
		      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               aTipos=""
		       aTipos0=""
		      If aTipos <> ")" Then
		        strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
		        strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and (Objprty.estado = 1 "
		        strsql = strsql & "Or Objprty.estado = 3)  " '&"And Objetos.tipobjeto In " & aTipos  & "Select nivel From Tipobjeto Where tipobjeto = '" & sOpcion & "'"
		        'response.write chr(13) & strsql & chr(13) & "<br>"
		        Set rsTipo2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
		      End If
		      If aTipos0 <> ")" Then
		        strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
		        strsql = strsql & "Tema.codtema = Prodserv.codtema " '&" and Prodserv.marcaps In " & aTipos0
		        'response.write strsql & "<br>"
		        Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
		      End If

		      'Esquema superior
		      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Esquema Superior", "si")
		      sInput = "<font class=at><select name=Esquema>"
		      If aTipos <> ")" Then
		         sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
		         sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
		         sInput = sInput & "<--          Seleccione un esquema          -->"
		         sInput = sInput & "</OPTION>"

		         Do Until rsTipo2.EOF

		            sInput = sInput & "<option value='" & rsTipo2(0) & "_" & rsTipo2(2) & "' "
		            If  bControl = "O" Then
		                If  Cstr(Nsect) = Cstr(rsTipo2(0)) Then
		                    sInput = sInput & " selected "
		                End If
		            End If
		            sInput = sInput & " >"

		            For inp = 0 To UBound(aevento)
		                If  aevento(inp) = UCase(rsTipo2(2)) Then
		                    sInput       = sInput & nevento(inp) & ": "
		                    Exit For
		                End If
		            Next

		            sInput = sInput & rsTipo2(1) & " - "

		            wwwtipo2 = rsTipo2(3)

		            if len(wwwtipo2 & "") > 40 then
		               sInput = sInput & left(wwwtipo2,40) & "..(+)"
		            else
		               sInput = sInput & wwwtipo2
		            end if


		      ''''      sInput = sInput & rsTipo2(3)

		            For inp = 0 To UBound(aestr)
		                If  aestr(inp) = UCase(rsTipo2(9)) Then
		                    sInput     = sInput & " (" & nestr(inp) & ")"
		                    Exit For
		                End If
		            Next

		            rsTipo2.movenext
		        Loop
		      End If

		      If aTipos0 <> ")" Then

		         sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
		         sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
		         sInput = sInput & "<--          Seleccione un subtema          -->"
		         sInput = sInput & "</OPTION>"


		         Do Until rsTipo0.EOF
		            sInput = sInput & "<option value='" & rsTipo0(1) & "_" & rsTipo0(5) & "' "
		            If  bControl = "S" Then
		                If  CStr(Nsect) = CStr(rsTipo0(1)) Then
		                    sInput  = sInput & " selected "
		                End If
		            End If
		            sInput = sInput & " >"

		            For inp = 0 To UBound(asector)
		                If  asector(inp) = UCase(rsTipo0(5)) Then
		                    sInput  = sInput & nsector(inp) & ": "
		                    Exit For
		                End If
		            Next

		            sInput = sInput & rsTipo0(0)
		            sInput = sInput & "/" & rsTipo0(2) & "/"

		            rsTipo0.movenext
		        Loop
		      End If

		      sInput = sInput & "</select></font>"
		      Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
		      Fila   = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

		      Tabla  = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")

		      If  aTipos  <> ")"  Then Call OFV.cerrarconsulta(rsTipo2)
		      If  aTipos0 <> ")"  Then Call OFV.cerrarconsulta(rsTipo0)

		      Tabla      = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")
		      Fila       = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

		      sInput     = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "this.style.display='none';")
		      Celda      = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
		      sAccion    = "Copiar2.asp?Codigo=" & nCodigo & "&Tipo=" & sTipo & "&Perfil=" & sPerfil & "&empresa=" & sEmpresa
		      sAccion    = sAccion & "&Opcion=" & sOpcion & "&Copiar=SI" & "&moveresquema=" & moveresquema & "&ESQSUP=" & ESQSUP
		      Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

		      sFunc      = "window.close()"
		      sInput     = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
		      Celda      = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
		      Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
		      Fila       = Formulario & OFV.GenFRow()

		      Tabla      = Fila & OFV.GenFTabla()

		      sHTML      = sHTML & "<br><br><br><br><center>" & Tabla & "</center>"
		      sHTML      = sHTML & "</body></html>"

      Call OFV.cerrarconsulta(rsObj)


 case else ' SOLO ESQUEMA


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
       IF  rsObj.EOF THEN
        XTipo=sTipo
       ELSE
       XTipo  = rsObj(2)
       END IF
       strsql = "Select codobjmaster,tipobjmaster From relaobj Where codojasoc =" & nCodigo
       strsql = strsql & " and tipobjasoc = '" & XTipo	 & "' "
      ' response.write strsql

       Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)



       bSector = "S/D"
       bControl = "N"
       Nsect = 0
       If Not rsNew.EOF Then
       Nsect = rsNew(0)
       For inp = 0 To UBound(asector)
           If asector(inp) = UCase(rsNew(1)) Then
              bSector  = nsector(inp)
              bControl = "S"
              Exit For
            End If
       Next

       If bControl = "N" Then
          For inp = 0 To UBound(aevento)
           If aevento(inp) = UCase(rsNew(1)) Then
              bSector  = nevento(inp)
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
          strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
          strsql = strsql & "Tema.codtema = Prodserv.codtema and prodserv.codprodserv = " & rsNew(0)
          Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
          If Not rsNew.EOF Then
             bSector = bSector & ": " & rsNew(0) & "/" & rsNew(2)
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


'''sEmpresa

        if  sEmpresa = "" then sEmpresa = 1
        bempresa = sEmpresa
        sLink    = "&nbsp;"

        Fila = ""

  IF  Session("EMPMUN") = "S" THEN

        FOR IAEMP=1 TO TAEMP

            IF ATREMP(IAEMP,2) = 1 THEN
               IF ATREMP(IAEMP,0) = sEmpresa THEN
                  bempresa = ATREMP(IAEMP,4)
                  sLink    = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
                  exit for
               END IF
            END IF

        NEXT

       'empresa
        Celda = OFV.GenCelda("", "ttt ", "", "", "", "", "", sLink, "si")
        sInput = "&nbsp;&nbsp;" & bempresa
        Celda = Celda & OFV.GenCelda("", "ttt", "", "", "", "", "", sInput, "si")
        Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

        Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;&nbsp;", "si")
        Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

   end if

       'origen
        Celda = OFV.GenCelda("", "tt", "", "", "", "", "", " Copiar:" , "si")
        sInput = bEvento & ": " & rsObj(1) & " - " & rsObj(3)
        For inp = 0 To UBound(aestr)
            If aestr(inp) = UCase(rsObj(9)) Then
               sInput = sInput & " (" & nestr(inp) & ")"
               Exit For
            End If
        Next
        Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




       'Superior
        Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Perteneciente a:", "si")
        sInput = bSector
        Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
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
        Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'Descripcion
      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Descripcion", "si")
      sInput = OFV.GenerarInput(rsObj(3).Name & sREADONLY, rsObj(3), "agregar", "Copiar", "80", "dt", "")
      Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'Identificacion
      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Identificacion", "si")
      sInput = OFV.GenerarInput(rsObj(1).Name & sREADONLY, rsObj(1), "agregar", "copiar", "20", "dt", "")
      Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      strsql = "Select nivel From Tipobjeto Where tipobjeto = '" & sOpcion & "'"
      Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
      sTipo = rsTipo(0)
      Call OFV.cerrarconsulta(rsTipo)
      Select Case sTipo
      Case "C"
          strsql = "Select tipobjeto From Subtema Where levc = 1"
      Case "D"
          strsql = "Select tipobjeto From Subtema Where levd = 1"
      Case "E"
          strsql = "Select tipobjeto From Subtema Where leve = 1"
      Case "F"
          strsql = "Select tipobjeto From Subtema Where levf = 1"
      Case "G"
          strsql = "Select tipobjeto From Subtema Where levg = 1"
      Case "H"
          strsql = "Select tipobjeto From Subtema Where levh = 1"
      Case "I"
          strsql = "Select tipobjeto From Subtema Where levi = 1"
      Case "J"
          strsql = "Select tipobjeto From Subtema Where levj = 1"
      Case "K"
          strsql = "Select tipobjeto From Subtema Where levk = 1"
      Case "L"
          strsql = "Select tipobjeto From Subtema Where levl = 1"
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
          strsql = "Select tipobjeto From Tipobjeto Where levc = 1"
      Case "D"
          strsql = "Select tipobjeto From Tipobjeto Where levd = 1"
      Case "E"
          strsql = "Select tipobjeto From Tipobjeto Where leve = 1"
      Case "F"
          strsql = "Select tipobjeto From Tipobjeto Where levf = 1"
      Case "G"
          strsql = "Select tipobjeto From Tipobjeto Where levg = 1"
      Case "H"
          strsql = "Select tipobjeto From Tipobjeto Where levh = 1"
      Case "I"
          strsql = "Select tipobjeto From Tipobjeto Where levi = 1"
      Case "J"
          strsql = "Select tipobjeto From Tipobjeto Where levj = 1"
      Case "K"
          strsql = "Select tipobjeto From Tipobjeto Where levk = 1"
      Case "L"
          strsql = "Select tipobjeto From Tipobjeto Where levl = 1"
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
        strsql = strsql & "Or Objprty.estado = 3) And Objetos.tipobjeto In " & aTipos & "and Objetos.aplic in ('EM', 'PL') "
      '  response.write strsql & "<br>"
        Set rsTipo2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
      End If
      If aTipos0 <> ")" Then
        strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
        strsql = strsql & "Tema.codtema = Prodserv.codtema and Prodserv.marcaps In " & aTipos0
        'response.write strsql & "<br>"
        Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
      End If

      'Esquema superior
      Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Esquema Superior", "si")
      sInput = "<font class=at><select name=Esquema>"
      If aTipos <> ")" Then
         sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
         sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
         sInput = sInput & "<--          Seleccione un esquema          -->"
         sInput = sInput & "</OPTION>"

         Do Until rsTipo2.EOF

            sInput = sInput & "<option value='" & rsTipo2(0) & "_" & rsTipo2(2) & "' "
            If  bControl = "O" Then
                If  Cstr(Nsect) = Cstr(rsTipo2(0)) Then
                    sInput = sInput & " selected "
                End If
            End If
            sInput = sInput & " >"

            For inp = 0 To UBound(aevento)
                If  aevento(inp) = UCase(rsTipo2(2)) Then
                    sInput       = sInput & nevento(inp) & ": "
                    Exit For
                End If
            Next

            sInput = sInput & rsTipo2(1) & " - "

            wwwtipo2 = rsTipo2(3)

            if len(wwwtipo2 & "") > 40 then
               sInput = sInput & left(wwwtipo2,40) & "..(+)"
            else
               sInput = sInput & wwwtipo2
            end if


      ''''      sInput = sInput & rsTipo2(3)

            For inp = 0 To UBound(aestr)
                If  aestr(inp) = UCase(rsTipo2(9)) Then
                    sInput     = sInput & " (" & nestr(inp) & ")"
                    Exit For
                End If
            Next

            rsTipo2.movenext
        Loop
      End If

      If aTipos0 <> ")" Then

         sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
         sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
         sInput = sInput & "<--          Seleccione un Subtema          -->"
         sInput = sInput & "</OPTION>"


         Do Until rsTipo0.EOF
            sInput = sInput & "<option value='" & rsTipo0(1) & "_" & rsTipo0(5) & "' "
            If  bControl = "S" Then
                If  CStr(Nsect) = CStr(rsTipo0(1)) Then
                    sInput  = sInput & " selected "
                End If
            End If
            sInput = sInput & " >"

            For inp = 0 To UBound(asector)
                If  asector(inp) = UCase(rsTipo0(5)) Then
                    sInput  = sInput & nsector(inp) & ": "
                    Exit For
                End If
            Next

            sInput = sInput & rsTipo0(0)
            sInput = sInput & "/" & rsTipo0(2) & "/"

            rsTipo0.movenext
        Loop
      End If

      sInput = sInput & "</select></font>"
      Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
      Fila   = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      Tabla  = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")

      If  aTipos  <> ")"  Then Call OFV.cerrarconsulta(rsTipo2)
      If  aTipos0 <> ")"  Then Call OFV.cerrarconsulta(rsTipo0)

      Tabla      = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")
      Fila       = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

      sInput     = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "this.style.display='none';")
      Celda      = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
      sAccion    = "Copiar2.asp?Codigo=" & nCodigo & "&Tipo=" & sTipo & "&Perfil=" & sPerfil & "&empresa=" & sEmpresa
      sAccion    = sAccion & "&Opcion=" & sOpcion & "&Copiar=SI" & "&moveresquema=" & moveresquema & "&ESQSUP=" & ESQSUP
      Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

      sFunc      = "window.close()"
      sInput     = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
      Celda      = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
      Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
      Fila       = Formulario & OFV.GenFRow()

      Tabla      = Fila & OFV.GenFTabla()

      sHTML      = sHTML & "<br><br><br><br><center>" & Tabla & "</center>"
      sHTML      = sHTML & "</body></html>"

      Call OFV.cerrarconsulta(rsObj)


  end select ' esqsup

   Call OFV.cerrarconn(cn)

    End If








end if

response.write sHTML

%>