<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	Pantalla=request.querystring("Pantalla")
	session("ModMain") = "CL"
	wmnnoved = "NOVED" & WKAPLIC

	session(wmnnoved) = "SI"
	aplcheck()

	wspviewopc = 2

	stilo1 = " style={background-color:" & session("CMFMENU") & ";border-style:outset;border-width:2;}"
	stilo2 = " style={background-color:" & session("CMFMNBOT") & ";}"


	wmsgs = request.querystring("xxmsg")


	ReDim Preserve bConsultar(10)

	''''response.write "<br><br><br><br>" & Session("Form") & "<br>"
	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	If  session("APLMDCLI")="BC" Then
		xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If
	xs = xs & " myWin= open(Datos"

	xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," '& Chr(34)
	' xs = xs & "'menubar=no,toolbar=no,height=h,width=w,resizable=yes'" '& Chr(34)
	xs = xs &"  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'" '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)



	Novedform = "CLIMN16317206150930885546227"

	If (session("UN") & "") = "" Then


		XXXX = NEWEMPRE("NO")


		SEMPRE = ATREMP(0,0)
		session("UN") = SEMPRE & ""

		If (session("UN") & "") = "" Then session("UN") = 1


	End If


	Set cn = ofv.conectar(ofv.strconn2)
	Set cn0 = ofv.conectar(ofv.strconn0)

	wmnfuncion = ""

	atmensajes = 0
	Dim atitsecciones(50,5)
	ttitsecc = 0
	acctitblks = 0
	acctitlin = 0

	Dim atit(50)
	Dim awmnasp(50)
	wmntope = 10
	wmntopeX = 10
	WMNINITX = 2

	Dim ANTabla(500,25)
	TNTabla = 0
	INTabla = 0

	For TT=1 To wmntope
		atit(tt) = "NO"
	Next

	titc = 0
	Strsql = "SELECT * FROM novedtit "
	Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rsx.eof Then

		For tt = WMNINITX To wmntopeX
			atit(tt) = rsx(tt)
			If rsx(tt) <> "NO" Then
				OFV.ObtenerAtributosNOdo Novedform, "", "TIPO", CStr(tt)
				bConsultar(TT) = OFV.AConsultar

				If TT = 10 Then

					If bConsultar(TT) Then
						session("SMPLMENU") = "N"
						atit(tt) = "NO"

					Else
						If session("SMPLMENU") <> "S" Then
							session("SMPLMENU") = "S"
							wmnfuncion = "<script language='javascript' > hdemenu(); </script> "
						End If
						titc = titc + 1



					End If


				Else

					If bConsultar(TT) Then
						titc = titc + 1

					Else

						atit(tt) = "NO"

					End If

				End If

			End If
		Next

		If  atit(10) <> "NO" Then

			ii = 10
			'  xxxx = wmnbuscamenu("CLIS150172171122331051096140")
			wmntope = ii

		End If

	End If
	Call ofv.cerrarconsulta(rsx)

	menuSel = 2
	menu    = 2

	If menu = "" Then

		If   atit(3) <> "NO" Then
			'menu = 3
			If WKAPLIC = "OR" Then
				menu = 4
				menuSel = 4
			Else
				menu=3
				menuSel = 3
			End If


		Else
			menu = 0
			menuSel = 0

			For iix = 1 To 10
				If   atit(iix) <> "NO" Then
					menu = iix
					menuSel = iix
					Exit For
				End If
			Next
		End If

	End If

	menu    = CInt(menu)

	sHTML = ""



	TABLA1 = ""
	TABLAB = ""
	TABLAN = ""
	TABLAB1 = ""




	nFecini = 0

	If nFecini < CLng(CDate("01-01-1900")) Then
		nFecini = CLng(CDate("01-01-1900"))
	Else
		nFecini = 0
	End If

	dFechai  = CLng(Date()) - nFecini
	XTrColor = "#FFFFFF"

	sglmar   = "<font class=btiW >*</font>"
	sglman   = "<font class=btiW >1</font>"
	sglprf   = "<font class=btic >&#0128;</font>"
	sgllnk   = "<font class=btic >&#0126;</font>"
	sgladj   = "<font class=btic >&#0204;</font>"




	sHTML = "<html><body onload='window.top.ldfobj();' >"

	xxxx = noveselec(2,wmsgs)

	''''''response.write ANTabla(INTabla,10) & "-11111<br>"
	TNTabla = INTabla
	xxxx = noveformat(2,2)

	'' sHTML = sHTML &  "<h4>comienzo</h4>" & ttitsecc & "-" & wmsgs

	sHTML = sHTML &  atitsecciones(1,3) & xs

	''  sHTML = sHTML &  "<h4>fin</h4>"

	sHTML = sHTML &  "</body></html>"



	response.write sHTML
End If




Function fechaArg(f)
	dia=Day(f)
	mes=Month(f)
	ano=Year(f)
	fechaarg=dia &"/" & mes & "/" & ano
End Function



Function noveselec(xmenu,nnmsg)

	'' xscolor = "#999999 size=2 "

	xscolor = " class=istt "'"#0000 style='font-size:8pt;font-weight:normal;' "

	Do Until SaleSel = "S"

		'response.write "<br> XMENU: "& xmenu & " NNMSG: " & nnmsg &"<br>"

		Select Case xmenu


			Case 1


			Case 2

			If nnmsg <> "" Then

				' incidennte 1377
				Strsql = "SELECT distinct objid,codigo,tipo,fecinicio,fecexp,mensaje,importancia,relacion,codobjeto,tipobjeto,workflowid,contacto,hora,msgversion,msgtitulo,msgcopete,msgimagen,msgurl,msgurldsc,msgvaenc,msgvaforo,msgcodenc,msgcodforo FROM csnovedades "
				'Strsql = Strsql & " Where codigo = '" & nnmsg & "' "
				Strsql = Strsql & " Where (usugrp='TODOS' OR usugrp='"  & session("usuario") &"' or usugrp='CORP') AND codigo = '" & nnmsg & "' "

				SaleSel = "S"

			Else


				Strsql = "SELECT * FROM csnovedades "
				Strsql = Strsql & " Where fecexp > " & dFechai
				Strsql = Strsql & " and fecinicio < " & dFechai + 1
				Strsql = Strsql & " And ("
				Strsql = Strsql & " (tipo = 2 And usugrp = '" & session("usuario") & "') or "
				Strsql = Strsql & " (tipo = 1 And usugrp = 'TODOS' and sector = 999 )    or "
				Strsql = Strsql & " (tipo = 1 And usugrp = 'TODOS' and sector = 0" & session("UN") & ") OR "
				Strsql = Strsql & " (tipo = 51 And usugrp = 'TODOS' and sector = 999 )    or "
				Strsql = Strsql & " (tipo = 51 And usugrp = 'TODOS' and sector = 0" & session("UN") & ")"
				Strsql = Strsql & ") "
				Strsql = Strsql & " order by fecinicio desc, tipo desc, importancia, objid desc"

				'''    session("ModMain") = "CL"




			End If
			'     RESPONSE.WRITE STRSQL & chr(13) & "<BR>"
			Set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If   rsx.eof Then
				xnohay1 = "S"
			Else
				xnohay1 = "N"
				'  do until rsx.eof
				INTabla = INTabla + 1

				xfromusr = rsx(11)

				xxxx = getInfoUser(rsx(11))

				xfromusr = session("SGU_nombre")

				ANTabla(INTabla,0) = rsx(1)
				ANTabla(INTabla,1) = CLng(rsx(2))
				ANTabla(INTabla,2) = CLng(rsx(6))
				ANTabla(INTabla,3) = CLng(rsx(3))
				ANTabla(INTabla,4) = rsx(12)
				ANTabla(INTabla,5) = "N"
				ANTabla(INTabla,6) = "N"
				ANTabla(INTabla,7) = xfromusr
				ANTabla(INTabla,8) = rsx(5)
				ANTabla(INTabla,9) = "&nbsp;"
				ANTabla(INTabla,10) = "XXMSGXX"
				ANTabla(INTabla,11) = rsx(2)
				ANTabla(INTabla,12) = rsx(13)
				ANTabla(INTabla,13) = rsx(14)
				ANTabla(INTabla,14) = rsx(15)
				ANTabla(INTabla,15) = rsx(16)
				ANTabla(INTabla,16) = rsx(18)
				ANTabla(INTabla,17) = rsx(19) & ""
				ANTabla(INTabla,18) = rsx(20) & ""
				ANTabla(INTabla,19) = rsx(21) & ""
				ANTabla(INTabla,20) = rsx(22) & ""

				ANTabla(INTabla,21) = "0"

				'   response.write ANTabla(INTabla,10) & "-11111<br>"

				' rsx.movenext
				' loop
			End If
			Call ofv.cerrarconsulta(rsx)

			Case 3,9

			'''    session("ModMain") = "CL"

			Strsql = "SELECT distinct * FROM csnovedades "
			Strsql = Strsql & " Where fecexp > " & dFechai
			Strsql = Strsql & " and fecinicio <  " & dFechai + 1
			Strsql = Strsql & " And (tipo = 3 or tipo = 9)"
			Strsql = Strsql & " And ("
			Strsql = Strsql & "    (usugrp = '" & session("usuario") & "') "
			Strsql = Strsql & " or ( usugrp = 'TODOS' and sector = 999)"
			Strsql = Strsql & " or ( usugrp = 'TODOS' and sector = 0" & session("UN") & ")"
			Strsql = Strsql & "     ) "
			Strsql = Strsql & " order by fecinicio desc, importancia, objid desc"
			'   RESPONSE.WRITE STRSQL & chr(13) & "<BR>"
			Set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If   rsx.eof Then
				xnohay1 = "S"

			Else

				xvinoant = "no"
				Do Until rsx.eof


					If xvinoant <> rsx(1) Then
						xvinoant = rsx(1)

						xvale = "N"

						Strsql  = "SELECT * FROM objetos "
						Strsql  = Strsql & " Where codobjeto = 0" & rsx(8)
						Set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsxp.eof Then
							xvale = "S"
						End If
						Call ofv.cerrarconsulta(rsxp)

						If  xvale = "S" Then

							xnohay1 = "N"
							INTabla = INTabla + 1


							xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


							sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(8) & "&Tipo=" & rsx(9)
							sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
							sHref = "&nbsp;"
							' response.write sActcon
							sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)


							wmsgx = rsx(5) & ""

							If Len(wmsgx) < 5 Then
								wmsgx = esqnomP(rsx(9),rsx(8),ofv) & " VIGENTE A PARTIR DEL " & fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900")))) & ". "
							Else
								'                                  response.write "aa" & len(wmsgx) & rsx(5)
								wmsgx = UCase(wmsgx)
							End If

							wNomL = wmsgx & "&nbsp;"



							sImagen = wNomL

							If xarc = "NO" Then
								xarc = "&nbsp;"
							End If

							sLink = sImagen
							'   XTSLNKM = OFV.GenLink("", "cmb title='Consultar Documento...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
							XTSLNKM = OFV.GenLink("", "istt title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

							xfromusr = rsx(11)


							xxxx = getInfoUser(rsx(11))

							xfromusr = session("SGU_nombre")

							ANTabla(INTabla,0) = rsx(1)
							ANTabla(INTabla,1) = rsx(2)
							ANTabla(INTabla,2) = CLng(rsx(6))
							ANTabla(INTabla,3) = CLng(rsx(3))
							ANTabla(INTabla,4) = rsx(12)
							ANTabla(INTabla,5) = "S"
							ANTabla(INTabla,6) = "N"
							ANTabla(INTabla,7) = xfromusr
							ANTabla(INTabla,8) = sLink
							ANTabla(INTabla,9) = xarc
							ANTabla(INTabla,10) = XTSLNKM
							ANTabla(INTabla,11) = xmenu

							If (rsx(13) & "") = "2" Then
								ANTabla(INTabla,12) = rsx(13)
								ANTabla(INTabla,13) = rsx(14)
								ANTabla(INTabla,14) = rsx(15)
								ANTabla(INTabla,15) = rsx(16)
								ANTabla(INTabla,16) = rsx(18)
							Else

								ANTabla(INTabla,12) = 2
								ANTabla(INTabla,13) = Capitalize(esqnomP(rsx(9),rsx(8),ofv))
								ANTabla(INTabla,14) = Capitalize(sLink)
								ANTabla(INTabla,15) = "0"
								ANTabla(INTabla,16) = ""
							End If

							ANTabla(INTabla,17) = rsx(19) & ""
							ANTabla(INTabla,18) = rsx(20) & ""
							ANTabla(INTabla,19) = rsx(21) & ""
							ANTabla(INTabla,20) = rsx(22) & ""

							ANTabla(INTabla,21) = rsx(8)

						End If
					End If

					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)

			Case 4

			'''     session("ModMain") = "WK"
			Session("Segce")   = "S"
			Session("Segco")   = "S"

			Strsql = "SELECT * FROM csnovedades "
			Strsql = Strsql & " Where fecexp > " & dFechai
			Strsql = Strsql & " and fecinicio < " & dFechai + 1
			Strsql = Strsql & " And tipo = 4"
			Strsql = Strsql & " And usugrp = '" & session("usuario") & "' order by fecinicio desc, importancia, objid desc"
			' RESPONSE.WRITE STRSQL
			Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  rsx.eof Then
				xnohay1 = "S"
			Else
				xnohay1 = "N"
				Do  Until rsx.eof


					INTabla = INTabla + 1

					Strsql = "SELECT archivo FROM tipobjeto "
					Strsql = Strsql & " Where tipobjeto = '" & rsx(9) & "'"
					Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If  Not rs.eof Then
						xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
					Else
						xarc = "NO"
					End If
					Call ofv.cerrarconsulta(rs)


					Session("ASP")      = "novedmenu.asp?menu=4"
					session("sEstado")  = "PARA APROBAR"
					session("sEstadon") = "2"
					Session("Form2")    = "apr"

					aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&Perfil=xxxx&Funcion=E"
					aspeditar = CStr(aspeditar)
					sHref     = "&nbsp;"
					' response.write sActcon
					sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

					wNomL = "&nbsp;" & UCase(rsx(5)) & "&nbsp;"

					sImagen = wNomL


					If xarc = "NO" Then
						xarc = "&nbsp;"
					End If

					''''  sLink = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=" & xscolor & " >" & sImagen & "</font>", "si")


					sLink = sImagen
					'XTSLNKM = OFV.GenLink("", "cmb ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
					XTSLNKM = OFV.GenLink("", "istt ", sHref, "", "", sFuncion, "XXMSGXX", "si")


					xfromusr = rsx(11)


					xxxx = getInfoUser(rsx(11))
					xfromusr = session("SGU_nombre")

					ANTabla(INTabla,0) = rsx(1)
					ANTabla(INTabla,1) = 4
					ANTabla(INTabla,2) = CLng(rsx(6))
					ANTabla(INTabla,3) = CLng(rsx(3))
					ANTabla(INTabla,4) = rsx(12)
					ANTabla(INTabla,5) = "S"
					ANTabla(INTabla,6) = "N"
					ANTabla(INTabla,7) = xfromusr
					ANTabla(INTabla,8) = sLink
					ANTabla(INTabla,9) = xarc
					ANTabla(INTabla,10) = XTSLNKM
					ANTabla(INTabla,11) = xmenu
					ANTabla(INTabla,12) = rsx(13)
					ANTabla(INTabla,13) = rsx(14)
					ANTabla(INTabla,14) = rsx(15)
					ANTabla(INTabla,15) = rsx(16)
					ANTabla(INTabla,16) = rsx(18)

					ANTabla(INTabla,17) = rsx(19) & ""
					ANTabla(INTabla,18) = rsx(20) & ""
					ANTabla(INTabla,19) = rsx(21) & ""
					ANTabla(INTabla,20) = rsx(22) & ""

					ANTabla(INTabla,21) = "0"

					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)

			'''''''''''''


			Set cn1 = ofv.conectar(ofv.strconn5)
			Set cn4 = ofv.conectar(ofv.strconn6)


			If wkaplic = "OR" Or wkaplic = "PL" Then

				xtipos = " OBJTIPO = 'RSK' "


			Else

				xtipos = " OBJTIPO in ("
				separ = ""

				Strsql = "SELECT tipobjeto FROM tipobjeto "
				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If  Not rs.eof Then
					Do While Not rs.eof
						xtipos = xtipos & separ & "'" & rs(0) & "'"
						separ = ","
						rs.movenext
					Loop
				Else
					xtipos = xtipos & "' '"
				End If
				xtipos = xtipos & ")"

			End If


			SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

			xxxx     = getInfoUser(Session("Usuario"))
			xusrlog  = session("SGU_nombre")



			If  xusrlog <> "NO" Then


				SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
				SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

				'''''''''''''''''''''''''''''''''''''''''''
				If  Session("DELLOG") = "S" Then

					XMFTGRP = Split(session("SegGrpLst"),Chr(9))
					xcount  = UBound(XMFTGRP) + 1
					XMFSEP  = ""
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

					For XMFI = 0 To UBound(XMFTGRP)
						SEGWHERE = SEGWHERE  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
						XMFSEP  = ","
					Next

					SEGWHERE = SEGWHERE & ") ) )"

				Else
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
				End If


				''''''''''''''''''''''''''''''''''''''''''

				'''   IF  RS(21) > 0 THEN


				If  CStr(session("SGU_idcostcenter") & "") <> "0" Then
					SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
					SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
					SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
					SEGWHERE = SEGWHERE & " ) ) "
				End If


			End If

			SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

			Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

			'RESPONSE.WRITE STRSQL & "<BR>" & chr(13)
			Set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
			If  rsx.eof Then
				xnohay2 = "S"
			Else

				Do  Until rsx.eof

					xvale = "N"

					If rsx(3) = "RSK" Then

						Strsql  = "SELECT * FROM ORRSKCIRCUITO "
						Strsql  = Strsql & " Where OBJID = 0" & rsx(4)
						Set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsxp.eof Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)


					Else



						Strsql  = "SELECT * FROM objetos "
						Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
						Set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsxp.eof Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

					End If




					If  xvale = "S" Then

						xnohay2 = "N"

						INTabla = INTabla + 1


						xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

						Session("ASP")      = "novedmenu.asp?menu=4"
						session("sEstado")  = "PARA APROBAR"
						session("sEstadon") = "2"
						Session("Form2")    = "apr"

						aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & RSX(8) & "&MAINAPL=" & xMAINAPL


						''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

						aspeditar = CStr(aspeditar)
						sHref     = "&nbsp;"
						'        response.write sActcon
						sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

						wNomL     = "" & RSX(7) & "&nbsp;-&nbsp;PASO&nbsp;"  & RSX(16) & "&nbsp;"

						sImagen   = wNomL


						sImagen =  sImagen & "-&nbsp;" & EsqnomP(rsx(3),rsx(4),ofv)

						If xarc = "NO" Then
							xarc = "&nbsp;"
						End If

						sLink = sImagen
						'XTSLNKM = OFV.GenLink("", "cmb title='Autorizar documento ...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
						XTSLNKM = OFV.GenLink("", "istt title='Autorizar documento ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

						xfromusr = rsx(6)


						xxxx = getInfoUser(rsx(6))
						xfromusr = session("SGU_nombre")

						ANTabla(INTabla,0) = "CKT" & INTabla
						ANTabla(INTabla,1) = 5
						ANTabla(INTabla,2) = 0
						ANTabla(INTabla,3) = dFechai
						ANTabla(INTabla,4) = rsx(10)
						ANTabla(INTabla,5) = "S"
						ANTabla(INTabla,6) = "N"
						ANTabla(INTabla,7) = xfromusr
						ANTabla(INTabla,8) = sLink
						ANTabla(INTabla,9) = xarc
						ANTabla(INTabla,10) = XTSLNKM
						ANTabla(INTabla,11) = xmenu


						ANTabla(INTabla,12) = 2
						ANTabla(INTabla,13) = Capitalize(esqnomP(rsx(3),rsx(4),ofv))
						ANTabla(INTabla,14) = Capitalize(sLink)
						ANTabla(INTabla,15) = "0"
						ANTabla(INTabla,16) = ""

						ANTabla(INTabla,17) = ""
						ANTabla(INTabla,18) = ""
						ANTabla(INTabla,19) = ""
						ANTabla(INTabla,20) = ""

						ANTabla(INTabla,21) = "0"

					End If



					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)

			Case 5

			'''    session("ModMain") = "CL"

			cuenta = 0

			''''''''''''''''''''''''''''''''
			XWHERE = " and USUARIO = '" & session("usuario") & "' "
			XWHERE = XWHERE & " AND CODOBJETO > 0 "

			If wkaplic = "OR" Or wkaplic = "PL" Then

				XWHERE = XWHERE & " AND TIPOBJETO = 'RSK' "

			Else

				XWHERE = XWHERE & " AND TIPOBJETO <> 'RSK' "

			End If

			XORDER = " group by codobjeto ORDER BY fecha desc "


			strsql = ""
			strsql = strsql & "SELECT CODOBJETO, max(TIPOBJETO), max(finicio) as fecha  "
			strsql = strsql & " FROM OBJALOGDET B1"
			strsql = strsql & " WHERE TACC = 'CONSEN'  " & XWHERE

			strsql = strsql & XORDER

			' RESPONSE.WRITE STRSQL & "<BR>"

			Set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If   rsx.eof Then
				xnohay1 = "S"

			Else

				xvinoant = "no"
				Do Until rsx.eof


					xvale = "N"



					If wkaplic = "OR" Or wkaplic = "PL" Then


						Strsql  = "SELECT * FROM ORRSKCIRCUITO "
						Strsql  = Strsql & " Where OBJID = 0" & rsx(0)
						Set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsxp.eof Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

						WKSPTIPOOBJ = "RSK"


					Else

						Strsql  = "SELECT * FROM objetos "
						Strsql  = Strsql & " Where codobjeto = 0" & rsx(0)
						''    RESPONSE.WRITE STRSQL & "<BR>"
						Set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsxp.eof Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)


						WKSPTIPOOBJ = "VER"

					End If




					If  xvale = "S" Then


						cuenta = cuenta + 1

						xnohay1 = "N"
						INTabla = INTabla + 1


						xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


						sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
						sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
						sHref = "&nbsp;"
						' response.write sActcon
						''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

						sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


						sTipobjeto = EsqnomP(WKSPTIPOOBJ,RSX(0),ofv)

						wmsgx = sTipobjeto & ""
						wNomL = wmsgx & "&nbsp;"



						sImagen = wNomL

						If xarc = "NO" Then
							xarc = "&nbsp;"
						End If

						sLink = sImagen
						'XTSLNKM = OFV.GenLink("", "cmb title='Consultar Documento...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
						XTSLNKM = OFV.GenLink("", "nt title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

						xfromusr = "COREMANAGER"


						xxxx = getInfoUser(xfromusr)

						xfromusr = session("SGU_nombre")


						ANTabla(INTabla,0) = "0"
						ANTabla(INTabla,1) = 3
						ANTabla(INTabla,2) = 1
						ANTabla(INTabla,3) = dFechai
						ANTabla(INTabla,4) = 0
						ANTabla(INTabla,5) = "S"
						ANTabla(INTabla,6) = "N"
						ANTabla(INTabla,7) = xfromusr
						ANTabla(INTabla,8) = sLink
						ANTabla(INTabla,9) = xarc
						ANTabla(INTabla,10) = XTSLNKM
						ANTabla(INTabla,11) = xmenu
						ANTabla(INTabla,12) = "1"
						ANTabla(INTabla,13) = ""
						ANTabla(INTabla,14) = ""
						ANTabla(INTabla,15) = ""
						ANTabla(INTabla,16) = ""

						ANTabla(INTabla,17) = ""
						ANTabla(INTabla,18) = ""
						ANTabla(INTabla,19) = ""
						ANTabla(INTabla,20) = ""

						ANTabla(INTabla,21) = "0"

					End If

					If cuenta > 9 Then Exit Do


					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)



			Case 6

			'''    session("ModMain") = "WK"
			Session("Segce")   = "S"
			Session("Segco")   = "S"

			' RESPONSE.WRITE "6"

			Set cn1 = ofv.conectar(ofv.strconn5)
			Set cn4 = ofv.conectar(ofv.strconn6)

			xtipos = " OBJTIPO in ('RSK'"
			separ = ","
			Strsql = "SELECT tipobjeto FROM tipobjeto "
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  Not rs.eof Then
				Do While Not rs.eof
					xtipos = xtipos & separ & "'" & rs(0) & "'"
					separ = ","
					rs.movenext
				Loop
			Else
				xtipos = xtipos & "' '"
			End If
			xtipos = xtipos & ")"

			SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

			xxxx     = getInfoUser(Session("Usuario"))
			xusrlog  = session("SGU_nombre")

			If  xusrlog <> "NO" Then


				SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
				SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

				'''''''''''''''''''''''''''''''''''''''''''
				If  Session("DELLOG") = "S" Then

					XMFTGRP = Split(session("SegGrpLst"),Chr(9))
					xcount  = UBound(XMFTGRP) + 1
					XMFSEP  = ""
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

					For XMFI = 0 To UBound(XMFTGRP)
						SEGWHERE = SEGWHERE  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
						XMFSEP  = ","
					Next

					SEGWHERE = SEGWHERE & ") ) )"

				Else
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
				End If


				''''''''''''''''''''''''''''''''''''''''''

				'''   IF  RS(21) > 0 THEN


				If  CStr(session("SGU_idcostcenter") & "") <> "0" Then
					SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
					SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
					SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
					SEGWHERE = SEGWHERE & " ) ) "
				End If


			End If

			SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

			Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

			'  RESPONSE.WRITE STRSQL & "4444<BR>"
			Set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
			If  rsx.eof Then
				xnohay2 = "S"
			Else

				Do  Until rsx.eof

					xvale = "N"

					'   Strsql  = "SELECT * FROM objetos "
					'   Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
					'   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					'   if  not rsxp.eof then
					'       xvale = "S"
					'   end if
					'   call ofv.cerrarconsulta(rsxp)

					If rsx(3) = "RSK" Then

						Strsql  = "SELECT * FROM ORRSKCIRCUITO "
						Strsql  = Strsql & " Where OBJID = 0" & rsx(4)
						Set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsxp.eof Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)


					Else



						Strsql  = "SELECT * FROM objetos "
						Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
						Set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsxp.eof Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

					End If


					If  xvale = "S" Then

						xnohay2 = "N"

						INTabla = INTabla + 1

						Strsql  = "SELECT archivo FROM tipobjeto "
						Strsql  = Strsql & " Where tipobjeto = '" & rsx(3) & "'"
						Set rs  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rs.eof Then
							'         xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
							'      else
							'          xarc = "NO"
						End If
						Call ofv.cerrarconsulta(rs)

						xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

						Session("ASP")      = "novedmenu.asp?menu=4"
						session("sEstado")  = "PARA APROBAR"
						session("sEstadon") = "2"
						Session("Form2")    = "apr"

						aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & RSX(8)


						''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

						aspeditar = CStr(aspeditar)
						sHref     = "&nbsp;"
						'        response.write sActcon
						sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

						wNomL     = "" & RSX(7) & "&nbsp;-&nbsp;PASO&nbsp;"  & RSX(16) & "&nbsp;"

						sImagen   = wNomL


						sImagen =  sImagen & "-&nbsp;" & EsqnomP(rsx(3),rsx(4),ofv)


						If xarc = "NO" Then
							xarc = "&nbsp;"
						End If

						sLink = sImagen
						XTSLNKM = OFV.GenLink("", "istt title='Ver mensaje ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")


						xfromusr = rsx(6)


						xxxx = getInfoUser(rsx(6))
						xfromusr = session("SGU_nombre")


						'' ANTabla(INTabla,0) = "0"
						ANTabla(INTabla,0) = "MISPEN" & INTabla
						ANTabla(INTabla,1) = 5
						ANTabla(INTabla,2) = 0
						ANTabla(INTabla,3) = dFechai
						ANTabla(INTabla,4) = rsx(10)
						ANTabla(INTabla,5) = "S"
						ANTabla(INTabla,6) = "N"
						ANTabla(INTabla,7) = xfromusr
						ANTabla(INTabla,8) = sLink
						ANTabla(INTabla,9) = xarc
						ANTabla(INTabla,10) = XTSLNKM
						ANTabla(INTabla,11) = xmenu
						ANTabla(INTabla,12) = 2
						ANTabla(INTabla,13) = Capitalize(esqnomP(rsx(3),rsx(4),ofv))
						ANTabla(INTabla,14) = Capitalize(sLink)
						ANTabla(INTabla,15) = "0"
						ANTabla(INTabla,16) = ""

						ANTabla(INTabla,17) = ""
						ANTabla(INTabla,18) = ""
						ANTabla(INTabla,19) = ""
						ANTabla(INTabla,20) = ""

						ANTabla(INTabla,21) = "0"

					End If



					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)





			Case 10

			'''   session("ModMain") = "CL"

			xnohay1 = "N"


			sLink = "<INPUT TYPE=BUTTON VALUE='HIDE MENU' onclick='hdemenu();' ><INPUT TYPE=BUTTON VALUE='SHOW MENU' onclick='mostrarmenu();' >"


			INTabla = INTabla + 1


			ANTabla(INTabla,0) = "0"
			ANTabla(INTabla,1) = 3
			ANTabla(INTabla,2) = 1
			ANTabla(INTabla,3) = dFechai
			ANTabla(INTabla,4) = 0
			ANTabla(INTabla,5) = "S"
			ANTabla(INTabla,6) = "N"
			ANTabla(INTabla,7) = "COREMANAGER"
			ANTabla(INTabla,8) = sLink
			ANTabla(INTabla,9) = "&nbsp;"
			ANTabla(INTabla,10) = "&nbsp;"
			ANTabla(INTabla,11) = xmenu
			ANTabla(INTabla,12) = "1"
			ANTabla(INTabla,13) = ""
			ANTabla(INTabla,14) = ""
			ANTabla(INTabla,15) = ""
			ANTabla(INTabla,16) = ""


			ANTabla(INTabla,17) = ""
			ANTabla(INTabla,18) = ""
			ANTabla(INTabla,19) = ""
			ANTabla(INTabla,20) = ""
			ANTabla(INTabla,21) = "0"








			Case Else


		End Select

		If  menuSel <> "" Then
			saleSel = "S"
		Else
			If  INTabla > 0 Then
				saleSel = "S"
			Else
				menuSel = 2
				menu    = 2
			End If
		End If



	Loop



End Function

Function noveformat(xmenu,XMENSEL)

	' response.write xmenu &  "<br>"
	' response.write xmensel &  "<br>"
	' response.write TNTabla &  "<br>"

	XTABLAN = ""

	acctitblks = 0
	acctitlin  = 0


	'xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "

	xsclase  = " CLASS=ietx "
	xsclase2 = "class=istt "
	'''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "

	XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='width:100%; "


	If XMENU = 2 Then

		'    XTABLAN = XTablaN & " width:40%;"


	Else

		'   XTABLAN = XTablaN & " width:30%;"

	End If



	XTABLAN = XTablaN & "border-style:solid;border-width:3;border-color:#ffffff;' >"


	XTABLAN = XTablaN &  "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
	XTABLAN = XTablaN &  "<tr HEIGHT=25  >"
	XTABLAN = XTablaN &  "<td " & xsclase & "   >" & Capitalize(atit(xmenu)) & "</TD>"
	XTABLAN = XTablaN &  "</TR>"
	XTABLAN = XTablaN &  "</table>"

	XTABLAN = XTablaN & "<div id='DTNOVINT" & xmenu & "' style='"

	XTABLAN = XTablaN & "border-style:solid;border-width:1;border-color:#ffffff;' >"

	XTABLAN = XTablaN &  "<table  width=99% cellspacing=0 cellpadding=0  >"


	''''  response.write TNTabla & "aaaa<br>"

	If  TNTabla = 0 Then


		XTABLAN = XTablaN &  "<tr>"
		XTABLAN = XTablaN &  "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(menu) & "&nbsp;</TD>"
		XTABLAN = XTablaN &  "</TR>"

		XTABLAN = XTablaN &  "<tr>"
		XTABLAN = XTablaN &  "<td class=ut colspan=8 >&nbsp;</TD>"
		XTABLAN = XTablaN &  "</TR>"


	Else

		'response.write TNTabla & "ssss<br>"


		For INTabla=1 To TNTabla

			'    acctitlin  = acctitlin + 2


			xtmcod   = ANTabla(INTabla,0)
			xtipom   = ANTabla(INTabla,1)
			xtpri    = ANTabla(INTabla,2)
			xtfvig   = ANTabla(INTabla,3)
			xthvig   = ANTabla(INTabla,4)
			xtsadj   = ANTabla(INTabla,5)
			xtslnk   = ANTabla(INTabla,6)
			xfromusr = ANTabla(INTabla,7)
			xtmsg    = ANTabla(INTabla,8)

			xtmsgimg    = ANTabla(INTabla,9)
			xtmsglnk    = ANTabla(INTabla,10)
			xtmsgTYPE   = ANTabla(INTabla,11)

			'response.write "MENSAJE:  "& xtmsg
			'response.write "INTTABLA: "& INTabla
			'response.write "PRIORIDAD:"&	xtpri
			'response.write "TIPOM:"&	xtipom
			xfclas    = "dt"
			sImage    = "&nbsp;"

			Select Case xtpri
				Case 0

				xfclas = "btbn"
				sglmod = 5
				sgltit = "Prioridad Alta."

				Case 1

				xfclas = "btba"
				sglmod = 4
				sgltit = "Prioridad Normal."

				Case 2

				xfclas = "btbs"
				sglmod = 6
				sgltit = "Prioridad baja."


			End Select

			sglmod     = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
			sImage     = ofv.GenLink("",xfclas & " title='" & sgltit & "' ","","","","",sglmod,"si")

			If  xtfvig = dFechai Then
				XMARQ  = sglmar
			Else
				XMARQ  = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
			End If

			If  xtipom = 5 Then
				xtfechavig = xthvig
				XTCOUNT    = InStr(xtfechavig," ")
				If XTCOUNT > 0 Then
					xtfechavig = Left(xtfechavig,XTCOUNT)
					xtfechavig = GTFechadorsh(CLng(CDate(xtfechavig)) - CLng(CDate("01-01-1900")))
				End If
			Else
				xtfechavig = GTFechadorsh(xtfvig)
				XTCOUNT    = InStr(xtfechavig," ")
				If XTCOUNT > 0 Then xtfechavig = Left(xtfechavig,XTCOUNT)


			End If

			Xnlnk      = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
			Xnadj      = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

			If xtsadj  = "S" Then  Xnadj = sgladj
			If xtslnk  = "S" Then  Xnlnk = sgllnk









			xtmsgVRS    = "0" & ANTabla(INTabla,12)
			xtmsgVRS    = CInt(xtmsgVRS)
			xtmsgTIT    = ANTabla(INTabla,13)
			'response.write "TITULO: "& xtmsgtit
			xtmsgCOP    = ANTabla(INTabla,14)
			'response.write " COPETE: " & xtmsgcop

			xtmsgIMG    = ANTabla(INTabla,15)
			xtmsgVID    = ANTabla(INTabla,16)


			xtmsgVAENC      = ANTabla(INTabla,17)
			xtmsgVAFORO     = ANTabla(INTabla,18)
			xtmsgENCCOD     = ANTabla(INTabla,19)
			xtmsgFOROCOD    = ANTabla(INTabla,20)

			xtmsgdoccod    = ANTabla(INTabla,21)

			' SI VRS ES TIPO 2

			If xtmsgVRS = 2 Then

				acctitlin  = acctitlin + 2

				'xtmsglnk    = replace(xtmsglnk,"XXMSGXX","<font STYLE='color:#990000;font-size:10pt;font-weight:bold;' >" & "" xtmsgTIT & "</font>")

				xtmsglnk    =  Replace(xtmsglnk,"XXMSGXX", xtmsgTIT )

				'  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
				'  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
				'  XTABLAN = XTablaN &  "</TR>"

				XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
				XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
				For yy=1 To 6
					XTABLAN = XTablaN &  "<td class=UT width='15%'  >&nbsp;</TD>"
				Next
				XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
				XTABLAN = XTablaN &  "</TR>"

				'********* TITULO DE LA NOVEDAD ***********

				If Pantalla<>"MAIN" Then

					XTABLAN = XTablaN &  "<tr HEIGHT='30' >"
					XTABLAN = XTablaN &  "<td "& xsclase2 &" colspan=8  >" & xtmsglnk &" MENU:" & "</TD>"
					XTABLAN = XTablaN &  "</TR>"

				End If

				''********* IMPRIME EL COPETE *******

				If Pantalla<>"MAIN" Then

					XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
					XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
					XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP
					XTABLAN = XTablaN &  "<tr  >"
					XTABLAN = XTablaN &  "<td class=ut  VALIGN='TOP' align='center' ><B><font color='#AA0000' size=2 >&#0149;</font></B></TD>"
					XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='text-align:justify;' >"
					XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
					XTABLAN = XTablaN &  "<td class=UT  >&nbsp;</TD>"
					XTABLAN = XTablaN &  "</TR>"

				End If

				'******************MENSAJE DE NOVEDAD **********
				XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
				XTABLAN = XTablaN &  "</TR>"
				XTABLAN = XTablaN &  "<tr HEIGHT='30' >"
				XTABLAN = XTablaN &  "<td  colspan=8  >" & Replace (Replace(xtmsg,Chr(13),"<br>")," "," ")  & "</TD>"
				XTABLAN = XTablaN &  "</TR>"


				'*******************MENSAJE DE NOVEDAD ***********
				If xtipom = "9" Then

					acctitlin  = acctitlin + 2


					XTMSGSEND = ""
					''  XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=1 >&#0124;</font></B>"
					XTMSGSEND = XTMSGSEND & "&nbsp;" & Capitalize(xtmsgCOP)


					XTABLAN = XTablaN &  "<tr  >"
					XTABLAN = XTablaN &  "<td class=ut  VALIGN='TOP' align='center' ><B>"
					XTABLAN = XTablaN &  "<font color='#AA0000' size=2 >&#0187;</font>"
					XTABLAN = XTablaN &  "<font color='#0000AA' size=1 >&#0124;</font>"
					XTABLAN = XTablaN &  "</B></TD>"
					XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
					XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
					XTABLAN = XTablaN &  "<td class=UT  >&nbsp;</TD>"
					XTABLAN = XTablaN &  "</TR>"


					XTMSGSEND = ""
					XTMSGSEND = XTMSGSEND & "<font color=#888888 style='font-weight:bold;' size=1 >recomendado por " & ANTabla(INTabla,7) & "</font>"

					XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
					XTABLAN = XTablaN &  "<td class=UT align='right' colspan=8 STYLE='background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
					XTABLAN = XTablaN &  "</TR>"

				End If


				'''''''''''''''''''''''''''''
				XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
				XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
				XTABLAN = XTablaN &  "</TR>"

				If Pantalla<>"MAIN" Then

					If xtmsgdoccod <> "0"  Then



						Strsql = "Select * From objinfo where codobjeto = 0" & xtmsgdoccod & " "
						Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not RS1.EOF Then

							xtmsgCOP    = rs1(6)
							xtmsgIMG    = rs1(7)
							xtmsgVID    = rs1(9)


						End If




						Strsql = "Select * From objinfoIMG  where codobjeto = 0" & xtmsgdoccod & " ORDER BY objSECUEN, OBJID"

						wspimgdat = 6

					Else

						Strsql = "Select * From NOVEDADESIMG  where NOVEDAD = '" & xtmcod & "' ORDER BY SECUEN, OBJID"
						wspimgdat = 5
					End If



					If (xtmsgIMG & "") = "1" Then



						Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not RS1.EOF Then



							XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
							XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

							XTABLAN = XTablaN &  "<table cellspacing=0 cellpadding=1  width=100% >"
							XTABLAN = XTablaN &  "<tr   >"


							XSFSTYLE= " STYLE='display:block;' "

							Do While Not rs1.eof

								acctitlin  = acctitlin + 10

								'' ********* IMAGEN

								'   If Pantalla<>"MAIN" THEN
								xCelda = ""
								xCelda = xCelda & "<div id='DIMGN" & INTabla & "' style='margin-top:1;margin-left:1;"
								xCelda = xCelda & "width:90%;overflow:auto;"
								xCelda = xCelda & ""
								xCelda = xCelda & "' >" & Chr(13)



								xCelda = xCelda & "<IMG style='border-style:solid;border-width:1;border-color:#888888;' SRC='" & RS1(wspimgdat) & "' >"

								xCelda = xCelda & "</div>" & Chr(13)

								XTABLAN = XTablaN &  "<td CLASS=UT align='center' " & XSFSTYLE & " width=100% >" & xCelda & "</TD>"
								XSFSTYLE= " STYLE='display:none;' "
								'
								'  END IF

								'' ********* IMAGEN


								RS1.MOVENEXT
							Loop

							XTABLAN = XTablaN &  "</TR>"
							XTABLAN = XTablaN &  "</table>"

							XTABLAN = XTablaN &  "</TD>"
							XTABLAN = XTablaN &  "</TR>"


							XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
							XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
							XTABLAN = XTablaN &  "</TR>"


						End If






					ElseIf (xtmsgIMG & "") = "2" Then

						acctitlin  = acctitlin + 10


						XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
						XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
						XTABLAN = XTablaN &  xtmsgVID
						'   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
						XTABLAN = XTablaN &  "</TD>"
						XTABLAN = XTablaN &  "</TR>"


						XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
						XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
						XTABLAN = XTablaN &  "</TR>"

					End If

					acctitlin  = acctitlin + 2





					''''' PREGUNTA SI VA ENCUESTA '''''''''''''''''''''''

					'    xtmsgVAENC      = ANTabla(INTabla,17)
					'    xtmsgVAFORO     = ANTabla(INTabla,18)
					'    xtmsgENCCOD     = ANTabla(INTabla,19)
					'    xtmsgFOROCOD    = ANTabla(INTabla,20)



					If xtmsgVAENC = "S" Then

						acctitlin  = acctitlin + 2

						Strsql = "Select * From emencuestas  where objid = 0" & xtmsgENCCOD
						Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not rs1.eof Then

							xtmsgTITENC    =  Capitalize(rs1(17) & "")
							xtmsgENCDER    = "S"


							If (rs1(2) & "") = "1" Then
								Strsql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & xtmsgENCCOD
								Strsql = Strsql & " Order By BLKSECUEN,objid DESC"
								Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								If Not rs1.eof Then
									sblktvoto      = rs1(4) & ""
									sblkCODIGO     = rs1(0) & ""
									If sblktvoto = "2" Then
										xtmsgENCDER    = "N"
										xtmsgTITENC    = Capitalize(rs1(3) & "")



										XTABLAN = XTablaN &  "<tr  >"
										XTABLAN = XTablaN &  "<td class=ut  VALIGN='TOP' align='center' >&nbsp;</TD>"
										XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='color:#000000;' align='center' ><br>"

										XTABLAN = XTablaN &  "<table width=99% cellspacing=0 cellpadding=1  style='background-color:#eeeeee;border-style:solid;border-width:1;border-color:#dddddd;' >"

										XTABLAN = XTablaN &  "<tr HEIGHT=20  >"
										XTABLAN = XTablaN &  "<td class=vtI COLSPAN=3  >Encuesta</TD>"
										XTABLAN = XTablaN &  "</TR>"


										XTABLAN = XTablaN &  "<tr HEIGHT=20  >"
										XTABLAN = XTablaN &  "<td class=utI width=85% style='font-size:8pt;font-weight:normal;' >" & xtmsgTITENC & "</TD>"

										WSPDIVN = "DENC" & xtmsgENCCOD

										Strsql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & xtmsgENCCOD
										Strsql2 = Strsql2 & " and BLKcodigo = 0" & sblkcodigo
										Strsql2 = Strsql2 & " Order By ITMSECUEN,objid DESC"

										Set rs2 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
										Do While (Not rs2.EOF)

											acctitlin  = acctitlin + 2


											WSPFRMN = "FENC" & xtmsgENCCOD & "IT" & rs2(0)
											WSPFRMA = "EMENCUESTASACT.ASP?TP=ES&OP=ACT&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
											WSPFRMR = "EMENCUESTASACT.ASP?TP=ES&OP=RES&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
											XTABLAN = XTablaN &  "<td class=ut width=5% align='center' >"
											WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"
											XTABLAN = XTablaN &  "<input type='button' style='width:45px;background-color:#cc0000;color:#ffffff;font-weight:bold;border-style:solid;border-width:1;border-color:#888888;' "
											XTABLAN = XTablaN &  "  value='&nbsp;" & rs2(4) & "&nbsp;' "
											XTABLAN = XTablaN &  "  onclick='"
											'  XTABLAN = XTablaN &  "alert("" voto:" & rs2(4) & " "");"
											'  XTABLAN = XTablaN &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
											XTABLAN = XTablaN &  "document." & WSPFRMN & ".submit();"
											XTABLAN = XTablaN &  "document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display="  & Chr(34) & "block" & Chr(34) & ";"
											XTABLAN = XTablaN &  "' >"
											XTABLAN = XTablaN &  "</TD>" & WSPFORME
											rs2.Movenext
										Loop

										acctitlin  = acctitlin + 2

										xCelda = ""
										xCelda = xCelda & "<div id='" & WSPDIVN & "' style='z-index:1;margin-top:1;margin-left:1;display:none;"
										xCelda = xCelda & "width:99%;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
										xCelda = xCelda & "background-color:#F9F9F9;"
										xCelda = xCelda & "' >" & Chr(13)
										xCelda = xCelda & "&nbsp;"
										xCelda = xCelda & "</div>" & Chr(13)


										XTABLAN = XTablaN &  "<td class=ut width=5% >&nbsp;</TD>"
										XTABLAN = XTablaN &  "</TR>"

										XTABLAN = XTablaN &  "<tr HEIGHT=20  >"
										XTABLAN = XTablaN &  "<td class=utI width=85%  >&nbsp;</TD>"
										XTABLAN = XTablaN &  "<td class=vt COLSPAN=2 ALIGN='CENTER' ><A CLASS=NT HREF='" & WSPFRMR & "' TARGET='fesquemas' "
										XTABLAN = XTablaN &  " ONCLICK='document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display="  & Chr(34) & "block" & Chr(34) & ";' "
										XTABLAN = XTablaN &  "><B>Resultados</B></TD>"
										XTABLAN = XTablaN &  "<td class=ut width=5% >&nbsp;</TD>"

										XTABLAN = XTablaN &  "</TR>"

										XTABLAN = XTablaN &  "<tr   >"
										XTABLAN = XTablaN &  "<td class=vtI COLSPAN=4 >&nbsp;</TD>"
										XTABLAN = XTablaN &  "</TR>"



										XTABLAN = XTablaN &  "</table>" & xCelda
										XTABLAN = XTablaN &  "</TD>"
										XTABLAN = XTablaN &  "<td class=ut   >&nbsp;</TD>"
										XTABLAN = XTablaN &  "</TR>"





									End If

								End If


							End If





						End If

					End If


					''''''''''''''''''''''''''''

					XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
					XTABLAN = XTablaN &  "<td class=UT colspan=8 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >&nbsp;</TD>"
					XTABLAN = XTablaN &  "</TR>"





				End If '' PANTALLA MAIN





			Else

				acctitlin  = acctitlin + 1


				xtmsg    = Replace(LTrim(xtmsg & "")," ","&nbsp;")
				xtmsglnk    = Replace(xtmsglnk,"XXMSGXX",Capitalize(xtmsg))


				If xtmcod = "0" Then

					XTMSGSEND = "<font color=#000000 >" & xtmsglnk & "</font>"

				Else

					XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
					XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
					XTMSGSEND = XTMSGSEND & "<font color=#000000 >&nbsp;" & xtmsglnk & "</font>"

				End If



				XTABLAN = XTablaN &  "<tr bgcolor=" & xTrColor & " HEIGHT='25' >"
				XTABLAN = XTablaN &  "<td class=" & xfclas & " align='center' >" & xtmsgimg & "</TD>"
				XTABLAN = XTablaN &  "<td class=uti colspan=6 STYLE='text-align:justify;' >" & XTMSGSEND & "</TD>"
				XTABLAN = XTablaN &  "<td class=UT  >&nbsp;</TD>"
				XTABLAN = XTablaN &  "</TR>"


				If xtmcod <> "0" Then

					' acctitlin  = acctitlin + 1


					XTABLAN = XTablaN &  "<tr bgcolor=" & xTrColor & " height=5 >"
					XTABLAN = XTablaN &  "<td class=ut colspan=8 valign=top align=right style='border-top-style:solid;border-top-width:1;border-top-color:#eeeeee;' >&nbsp;</TD>"
					XTABLAN = XTablaN &  "</TR>"

				End If


			End If


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#eeeeff"
			Else
				xTrColor = "#FFFFFF"
			End If







		Next


	End If

	XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
	XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
	XTABLAN = XTablaN &  "</TR>"

	XTABLAN = XTablaN &  "</table>"
	XTABLAN = XTablaN & "</DIV>"
	XTABLAN = XTablaN & "</DIV>"



	If  TNTabla = 0 Then
		noveformat = ""
	Else

		If xmenu = 4 Then  atmensajes = TNTabla

		ttitsecc = ttitsecc + 1

		atitsecciones(ttitsecc,0) = Capitalize(atit(xmenu))
		atitsecciones(ttitsecc,1) = TNTabla
		atitsecciones(ttitsecc,2) = acctitlin + 2
		atitsecciones(ttitsecc,3) = XTablaN
		atitsecciones(ttitsecc,4) = 0
		atitsecciones(ttitsecc,5) = 0

		noveformat = XTablaN
	End If


End Function


%>
