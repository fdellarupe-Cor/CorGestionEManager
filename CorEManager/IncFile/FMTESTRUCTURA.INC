
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GENFMTESTRUCTURA(FNCODOBJ,FNTIPOBJ)


	FNSHTML = ""
	''response.write id

	Fila  = ""

	FNSHTML = FNSHTML &  ofv.GenTabla("","","","100%  ","","0","","0","0","","","","no")

	wcantlin = 0

	Dim Inx
	Dim nrs(100)


	strsql = " SELECT * FROM CSESTDET "
	strsql = strsql & " where detcon = " & ID & "  "
	If marca <> "CA" Then
		strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
	End If
	strsql = strsql & " and not ( "
	strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
	strsql = strsql & " ) and stmodi = '' "
	strsql = strsql & " order by TSEC  "

	'response.write strsql
	''response.write marca & " 01"

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.CursorType = 3
	Rs.CacheSize = 100

	Rs.Open strsql, cn

	' ars = rs.GetRows()
	If Not rs.eof Then
		ars = rs.GetRows()
		''response.write id & "ok"

	Else
		campos = rs.fields.count - 1
		ReDim  ars(cint(campos),0)
		ars(39,0) = "N"
		wentro=1
	End If


	rs.Close
	Set rs = Nothing

	sant = "X"
	bseci = 0
	bsecix = 0
	ReDim preserve bsec(8)
	ReDim preserve bnum(8)
	ReDim preserve bnumG(8)
	For i=0 To 8
		bsec(i) = 0
		bnum(i) = 0
		bnumG(i) = 0
	Next

	inx = 0



	Do While wentro = 0



		For Inx = inx To UBound(ars,2)



			SCLAVE   = ars(0,inx)
			Stip     = ars(5,inx)
			sForm    = "edestrdet" & SCLAVE
			smodi    = ars(29,inx)
			stipmodi = ars(30,inx)

			If marca = "CA" Then
				Mostrar = "SI"
				If Stip = "G" Or  Stip = "H" Or  Stip = "I"  Or  Stip = "J" Then
					If smodiant And stipmodiant = "E" Then
						smodi = smodiant
						stipmodi = stipmodiant
					ElseIf Not smodi Then
						Mostrar = "NO"

					End If
				Else
					smodiant = ars(29,inx)
					stipmodiant = ars(30,inx)
				End If

			End If

			If Mostrar <> "" Then
				If Marca = "SI" Then
					If Mostrar = "NO" Then
						If trim(SID) = trim(sclave) Then
							MCtipo = Stip
							Mostrar = "SI"
						End If
					Else
						If Stip = MCtipo Or Stip < MCtipo Then
							Mostrar = "NO"
						End If
					End If
				End If
			End If


			pmk = "<form name="  & chr(34) & "T" & ars(0,inx)  & chr(34) &  " method=post >"
			pmk = pmk & "<input type=hidden value=a ></form>"

			sImage      = "&nbsp;"
			xfclas      = "dt"

			RCHEK = "N"
			Celdakk = ""

			If bVersion = "SI" And smodi And SESSION("SHWMOD") <> "N" Then




				If stipmodi = "M" Then
					xfclas = "btbs"
					xftitl = "Modificado."

					'''''''' sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>" & ars(0,inx)

					'sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
					sglmod = "<div class='" & xfclas & "' ><i class='las la-user-edit' ></i></div>"
					sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
					''''''para mostrar el item modificado

					Strsql = "Select * from CSESTDET where objid = 0" & ars(34,inx)
					Set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If Not rs3.eof Then
						RCHEK = "S"
					End If
					''response.write strsql

				ElseIf stipmodi = "A" Then
					xfclas = "btba"
					xftitl = "Agregado."
					'sglmod = "<div class='" & xfclas & "' ><font class=btic >6</font></div>"
					sglmod = "<div class='" & xfclas & "' ><i class='las la-user-plus' ></i></div>"
					sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
				ElseIf stipmodi = "E" Then
					xfclas = "btbn"
					xftitl = "Eliminado."
					'sglmod = "<div class='" & xfclas & "' ><font class=btic >5</font></div>"
					sglmod = "<div class='" & xfclas & "' ><i class='las la-user-minus' ></i></div>"
					sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
				End If
			End If




			supersec = "" & ars(38,inx)
       '     response.write supersec & " " & ars(0,inx) & "<br>"
			supersecarr = split(supersec,".")

			spsecn = ubound(supersecarr)


			sPoint = "T" & ars(0,inx)

			' estado observacion

			'''response.write nCodigo & " nCodigo " & chr(13)
			'''response.write sTipo   & " sTipo "  & chr(13)

			If Nestado=5  And origen <> "ED"  Then

				Set cnx = ofv.conectar(ofv.strconn5)
				strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
				''response.write strsql
				Set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
				If Not rss.eof Then
					idcircuito = rss("segckt")
					rsscondinx = rss(0)
				End If
				Call ofv.cerrarconsulta(rss)
				Call ofv.cerrarconn(cnx)

				'''response.write idcircuito & " idcircuito "
				'''response.write rsscondinx & " rsscondinx "

				strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codcircuito = '"& idcircuito &"'  and codinx1 = '" & rsscondinx & "' and "
				strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' and "
				strsql = strsql & " codusr = '" & session("usuario") & "' "
				strsql = strsql & " order by objid"
				'response.write strsql
				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

				sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
				sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & sclave  & "&ID2=" & ID

				xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
				sHref   = sAccion
				If rs.eof Then
					sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"

				Else
					sInput  = "<div style='border-style:solid;border-color:#DC143C;border-width:1;' >" & "OBS" & "</div>"

				End If
				''Celdaobs  = ofv.Gencelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
				''Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")
				Celdaobs= ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
				Celdaobs = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",Celdaobs,"si")

			End If


			Celda  = Celdaobs & ofv.GenCelda("","uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")

			If Marca  = "CA" Then
				If   RCHEK = "S" Then
					''        	 	Celdakk = Celdakk & ofv.GenCelda("","","","","","","","&nbsp;","si")
				End If
			End If

			'Secuencia
			Select Case stip
				Case "A"
				t = 7
				bseci = 0
				bsecix = 0
				bsec(1) = ars(4,inx)
				bnum(0) = bnum(0) + 1
				bnumG(0) = bnumG(0) + 1
				For i=bseci + 1 To 8
					bnum(i) = 0
					bnumG(i) = 0
				Next
				Case "B"
				t = 8
				bseci = 1
				bsecix = 1
				bsec(2) = ars(4,inx)
				bnum(1) = bnum(1) + 1
				bnumG(1) = bnumG(1) + 1
				For i=bseci + 1 To 8
					bnum(i) = 0
					bnumG(i) = 0
				Next
				Case "C"
				t = 9
				bseci = 2
				bsecix = 2
				bsec(3) = ars(4,inx)
				bnum(2) = bnum(2) + 1
				bnumG(2) = bnumG(2) + 1
				For i=bseci + 1 To 8
					bnum(i) = 0
					bnumG(i) = 0
				Next
				Case "D"
				t = 10
				bseci = 3
				bsecix = 3
				bsec(4) = ars(4,inx)
				bnum(3) = bnum(3) + 1
				bnumG(3) = bnumG(3) + 1
				For i=bseci + 1 To 8
					bnum(i) = 0
					bnumG(i) = 0
				Next
				Case "E"
				t = 11
				bseci = 4
				bsecix = 4
				bsec(5) = ars(4,inx)
				bnum(4) = bnum(4) + 1
				bnumG(4) = bnumG(4) + 1
				For i=bseci + 1 To 8
					bnum(i) = 0
					bnumG(i) = 0
				Next
				Case "F"
				t = 12
				bseci = 5
				bsecix = 5
				bsec(6) = ars(4,inx)
				bnum(5) = bnum(5) + 1
				bnumG(5) = bnumG(5) + 1
				For i=bseci + 1 To 8
					bnum(i) = 0
					bnumG(i) = 0
				Next
				Case "G"
				t = 13




				Select Case cstr(spsecn & "")
					Case "1"
					ta = 7
					bseci = 0
					Case "2"
					'ta = 8
					ta = 7
					bseci = 1
					Case "3"
					ta = 9
					bseci = 2
					Case "4"
					ta = 10
					bseci = 3
					Case "5"
					ta = 11
					bseci = 4
					Case "6"
					ta = 12
					bseci = 5
				End Select


				bnum(6) = bnum(6) + 1
				bnumG(6) = bnumG(6) + 1

				bsecix = bseci
				bsec(7) = ars(4,inx)

				Case "H"
				Select Case sant
					Case "A"
					bseci = 0
					Case "B"
					bseci = 1
					Case "C"
					bseci = 2
					Case "D"
					bseci = 3
					Case "E"
					bseci = 4
					Case "F"
					bseci = 5
				End Select
				bnum(7) = bnum(7) + 1
				bnumG(7) = bnumG(7) + 1
				bsecix = bseci + 1
				bsec(7) = ars(4,inx)
				Case "I"
				Select Case sant
					Case "A"
					bseci = 0
					Case "B"
					bseci = 1
					Case "C"
					bseci = 2
					Case "D"
					bseci = 3
					Case "E"
					bseci = 4
					Case "F"
					bseci = 5
				End Select
				bnum(8) = bnum(8) + 1
				bnumG(8) = bnumG(8) + 1
				bsecix = bseci + 1
				bsec(7) = ars(4,inx)
				Case "J"
				Select Case sant
					Case "A"
					bseci = 0
					Case "B"
					bseci = 1
					Case "C"
					bseci = 2
					Case "D"
					bseci = 3
					Case "E"
					bseci = 4
					Case "F"
					bseci = 5
				End Select
				bnum(8) = bnum(8) + 1
				bnumG(8) = bnumG(8) + 1
				bsecix = bseci + 1
				bsec(8) = ars(4,inx)

			End Select

			If stip <> "H" And stip <> "I"  And stip <> "J" Then

				For ti = 0 To 17
					Pdef(t,ti) = Pdef(t-7,ti)
				Next
				Pdef(t,18) = "SI"


				If ars(31,inx) = "NO" Then
					For ti = 11 To 28
						If (ars(ti,inx)&"") = "" Then
							sti = ""
						Else
							sti             = ars(ti,inx)
							Pdef(t,ti - 11) = sti
						End If



					Next
					Pdef(t,18) = "NO"

				End If
			End If

			BsecD = ""
			For i=1 To bsecix
				BsecD = BsecD & bsec(i) & "."
			Next





			'response.write strsql
			If Pdef(t,18) = "NO" Then
				tx = t
			Else
				tx = t - 7
			End If


			PNumCls = Pdef(tx,14)
			If PNumCls = "" Then PNumCls = "Simple"



			If stip <> "H" And stip <> "I"  And stip <> "J" Then
				PNumTip(tx) = Pdef(tx,13)
				If  PNumTip(tx) = "" Then PNumTip(tx) = "1"
			End If

			If stip = "G" Then
				If Pdef(tx,9) = "NO" Then
					MPvin = "NO"
				Else
					MPvin = "SI"
				End If
			End If

			' Anula la numeracion para este tipo o elemento
			If  PNumTip(tx) = "NO" Then PNumCls = PNumTip(tx)

			PFont  = Pdef(tx,0)
			PFsize = Pdef(tx,1)
			PFcol  = Pdef(tx,2)
			Pbgcol = Pdef(tx,3)
			PBorde = Pdef(tx,4)
			PFtipo = Pdef(tx,5)

			If marca = "CA" Then
				PFont  = "VERDANA"
				If smodi Then
					PFcol  = "#000080"
				Else
					''               PFcol  = "#aaaaaa"
					PFcol  = "#666666"
				End If
				PFsize = "8"
				Pbgcol = ""
				PBorde = ""
				PFtipo = ""
			End If


			PSang  = Pdef(tx,6)
			txi = tx
			If tx > 6 Then txi = tx - 7
			PEspa  = Pdef(txi,8)



			Palign = Pdef(tx,16)
			If Palign = "" Then Palign = "left"
			Ptalig = Pdef(tx,17)


			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'' arma span inicial ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			If Marca <> "SI" And Marca <> "CA" Then
				Celda = Celda & Ofv.Gencelda("","cc","","1%","","","",BsecD & ars(4,inx),"si")
			Else
				Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
				If Marca = "CA" Then
					If   RCHEK = "S" Then
						''             CeldaKK = CeldaKK & ofv.GenCelda("","","","","","","","&nbsp;","si")
					End If
				End If
			End If

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			txx = ta

			xmpvin = "NO"

			If stip = "G" And Pdef(txx,9) = "SI" Then
				xmpvin = "SI"
			End If

			If stip = "G" And MPvin  = "NO" Then
				xmpvin = "NO"
			End If

			xPNumS = "NO"


			If stip = "G" Then
				If Pnumer = "SI" Then
					Select Case Ucase(Pnumcls)
						Case "SIMPLE"
						xPnumS = "SI"
						Case "COMPUESTA"
						xPnumS = "SI"
					End Select

				End If
			End If



			If PInden = "SI" Then

				If  bsecix = 7 Then
					bspan = "aaa colspan=2"
				Else

					If stip <> "G" And stip <> "H" And stip <> "I"  And stip <> "J" Then
						wwcolsp =  10
					Else
						wwcolsp =  8
						If (xPNumS = "NO"  And stip = "G") Then
							wwcolsp = wwcolsp + 1
							''  Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a1","si")
						End If
						If (xmpvin  = "NO"  And stip = "G") Then
							wwcolsp = wwcolsp + 1
							''   Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a2","si")
						End If

					End If
					Bspan = "aaa colspan=" & wwcolsp - bsecix

				End If
				CeldaKKp = ""
				For i=1 To bsecix
					Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
					If Marca = "CA" Then
						If   RCHEK = "S" Then CeldaKKp = CeldaKKp & ofv.GenCelda("","","","","","","","&nbsp;","si")
					End If
				Next
			Else
				If stip <> "G" And stip <> "H" And stip <> "I"  And stip <> "J" Then
					Bspan = "aaa colspan=10"
				Else
					wwcolsp =  8
					If (xPNumS = "NO"  And stip = "G") Then
						wwcolsp = wwcolsp + 1
						''     Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a1","si")
					End If
					If (xmpvin  = "NO"  And stip = "G") Then
						wwcolsp = wwcolsp + 1
						''      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a2","si")
					End If

					Bspan = "aaa colspan=" & wwcolsp
					If stip <> "G" Then
						Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
						If Marca = "CA" Then
							If   RCHEK = "S" Then CeldaKKp = ofv.GenCelda("","","","","","","","&nbsp;","si")
						End If
					End If
				End If
			End If


			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			If  stip <> "H" And stip <> "I"  And stip <> "J" Then
				If PBorde <> "" And PBorde <> "NO" Then
					PBCol = ""
					If Ucase(PBorde) = "SOLID" Then
						PBCol = ";border-color:#000000;border-width:1"
					Else
						PBCol = ";border-color: #cccccc"
					End If
					bspan = bspan & " STYLE='border-style:" & PBorde & PBCol & ";' "
					'''   bspan = bspan & " style='border-style:inset;border-color: #eeeeee;' "
				End If
			End If

			PFtipoS = ""
			PFtipoF = ""
			If PFtipo <> "" Then
				For VT = 1 To len(PFtipo)
					Select Case mid(PFtipo,vt,1)
						Case "1"
						PFtipoS = PFtipoS & "<B>"
						PFtipoF = PFtipoF & "</B>"
						Case "2"
						PFtipoS = PFtipoS & "<I>"
						PFtipoF = PFtipoF & "</I>"
						Case "4"
						PFtipoS = PFtipoS & "<U>"
						PFtipoF = PFtipoF & "</U>"
						Case "8"
						PFtipoS = PFtipoS & "<S>"
						PFtipoF = PFtipoF & "</S>"
					End Select
				Next
			End If

			pStyle = PFtipoS & "<Font face='" & pfont & "' color='" & pfcol
			pStyle = pStyle & "' STYLE='font-size:" & pfsize & "pt;'>"
			pStyleF = "</Font>" & PFtipoF

			' TITULOS Y SUBTITULOS --------------------------------------------------

			If stip <> "G" And stip <> "H" And stip <> "I"  And stip <> "J" Then
				'Descripcion
				PNumS = ""
				If Pnumer = "SI" Then
					Select Case Ucase(Pnumcls)
						Case "SIMPLE"
						Select Case PnumTip(tx)
							Case "1"
							BnumS = BnumG(bseci)
							Case Else
							BnumS = ConvNumer(BnumG(bseci),PnumTip(tx))
						End Select

						PnumS = BnumS & ". "
						Case "COMPUESTA"
						For i=0 To bseci
							Select Case PnumTip(i + 7)
								Case "NO"
								BnumS = ""
								Case "1"
								BnumS = BnumG(i) & "."
								Case Else
								BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
							End Select

							PnumS = PnumS & BnumS
						Next
						PnumS = PnumS & " "
					End Select

				End If

				Celda = Celda & ofv.GenCelda(sPoint,bspan,Palign,"","",pbgcol,"","<div class=titulo> "  & pStyle & PnumS   & ars(7,inx)& "</div>" & pStyleF,"si")



				If Marca = "CA" Then

					If   RCHEK = "S" Then
						Filatt = ""
						Celdakk = ""
						Celdakk = Celdakk & ofv.GenCelda("","utg colspan=12 style='text-align:justify;color:#0054a9;' ","","30%","","","","<u>Version Anterior:</u>","si")
						Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

						Srs37VKK = rs3(7)
						Celdakk = ""
						Celdakk = Celdakkp
						''            CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & "dt valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"",Pnums  & Srs37VKK & "","si")
						CeldaKK = CeldaKK & ofv.GenCelda("",replace(bspan,"aaa","utg"),Ptalig & "utg valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"",Pnums & Srs37VKK & "","si")
						Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

					End If

				End If

				'   Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;x1","si")

				'''''''''' arma tabla de indices '''''''''''''''''''''''''
				stexto1 = sfidkB & "&nbsp;&nbsp;" & replace((PnumS & ars(7,inx) & "")," ","&nbsp;")

				For nn=0 To (bseci+1)
					stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
				Next

				Celdat = stexto1

				stexto1 = "<A class=nt href='javascript:void(0);' "
				''  stexto1 = stexto1 & " onclick=""creawin('T" & ars(0,inx) & "'); return false;"" "
				stexto1 = stexto1 & "   ><big>" & Celdat & "</big></a>"

				Celdat = "<td class=ut width=100% style='cursor:pointer;' onclick=""creawin('T" & ars(0,inx) & "');"" >" & stexto1 & "</td>"
				Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




			End If

			' PARRAFOS DE TEXTO    --------------------------------------------------

			If stip = "G" Then

				PNumS = ""
				If Pnumer = "SI"  Then
					Select Case Ucase(Pnumcls)
						Case "NO"
						BnumG(6) = BnumG(6) - 1
						' response.write "22222222222222222222"

						Case "SIMPLE"
						Select Case PnumTip(tx)
							Case "1"
							BnumS = BnumG(6)
							Case Else
							BnumS = ConvNumer(BnumG(6),PnumTip(tx))
						End Select

						PnumS = BnumS & ". "
						Case "COMPUESTA"
						For i=0 To 6
							If BnumG(i) > 0 Then
								Select Case PnumTip(i + 7)
									Case "NO"
									BnumS = ""
									Case "1"
									BnumS = BnumG(i) & "."
									Case Else
									BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
								End Select
								PnumS = PnumS & BnumS
							End If
						Next
						PnumS = PnumS & " "
					End Select

				End If

				tx = ta
				'response.write ta
				If  Ptalig = "" Then
					Ptalig    = Pdef(tx,17)
					If Ptalig = "" Then Ptalig = "left"
				End If

				If PSang = "" Or PSang = " " Then PSang = 0
				PSangS = ""

				For Vs = 1 To clng(PSang)
					PSangS = PSangS & "&nbsp;"
				Next

				If Pdef(tx,9) <> "SI" Then
					pVin = ""
				Else
					PVFont  = Pdef(tx,15)
					PVtipo  = Pdef(tx,10)
					PVind   = Pdef(tx,11)

					If PVind = "" Then PVind = 0

					PVindS = ""
					For Vi = 1 To pVind
						PVindS = PVindS & "&nbsp;"
					Next

					pVin = PFtipoS & "<Font face='" & PVFont & "' color='" & pfcol
					pVin = pVin & "' STYLE='font-size:" & pfsize & "pt;'>"
					pVin = pVin & PVtipo & "</font>" & PFtipoS & PVindS
				End If

				'Texto
				If MPVin <> "SI" Then
					pVin = ""
				End If

				If PnumS <> "" Then
					Celda = Celda & ofv.GenCelda("","aaa valign=TOP","","1%","","","",pStyle & PnumS  & pStyleF,"si")
				End If

				If PVin <> "" Then
					Celda = Celda & ofv.GenCelda("","aaa valign=TOP","","1%","","","",pVin ,"si")
				End If

				sTextoV =  ars(8,inx)


				' sTextoV =reemplazaspan(sTextoV)
				'    response.write "ssss" & len(sTextoV)






				''''''''''''''''''''


				'''''''''''''''''''

				sImagen  = "NO"
				wasotxt  = "N"


				'link relacionados al contenido
				descri=""
				wstextoLNK = ars(32,inx) & ""

				If Not isnull(wstextoLNK) And LEN(wstextoLNK) > 0 And wstextoLNK <> " " Then

					If wstextoLNK <> "0" Then

						wlink  = clng(wstextoLNK)
						strsql = "SELECT *  From objlnk where objid = " & wlink
						'response.write strsql
						Set rst = ofv.crearconsultaEx(StrSql,cn,1,volver)
						If Not rst.eof Then

							''''  wasotxt     = rst(14) & ""
							'''   wasotxtstr  = CLNG("0" & rst(15))

							''wasotxtcan  = CLNG("0" & rst(16)) 'columna linkasoc da error si no tiene datos'
							'wasotxtcan  = rst(16)
							'if wasotxtcan ="" then wasotxtcan="http://#" end if
							'''''wasotxtdat  = rst(17) & ""

							'  response.write rst(16).name

							wasotxt     = rst(14) & ""
							If isnumeric(rst(16))Then
								wasotxtcan  = rst(16)
							Else
								wasotxtcan  =0
							End If
							If isnumeric(rst(15))Then
								wasotxtstr  = CLNG("0" & rst(15))
							Else
								wasotxtstr=0
							End If
							wasotxtdat  = rst(17) & ""
							If wasotxt = "" Then wasotxt = "N"



							wNV1    = rst(11) & ""
							wNV2    = rst(12) & ""
							wNV3    = rst(13) & ""

							If  WNV1 = "" Then WNV1 = "0"
							If  WNV2 = "" Then WNV2 = "0"
							If  WNV3 = "" Then WNV3 = "0"

							codobjetoL = clng(rst("codobjetoL"))
							wtipoLc    = rst(7)
							If trim(rst("Ltipo")) = "L" Then

								wUrlDesc = rst(9) & ""


								wNomL    = wUrlDesc
								If wNomL = "" Then
									wNomL = "S/D"


									Strsql = "Select * From csGenObjetos where estado <> 4 and codobjeto = " & codobjetoL
									'response.write Strsql
									Set rst2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
									If Not rst2.eof Then
										wNomL = rst2(1)
									Else
										wNomL =""' "S/D"
									End If
									Call ofv.cerrarconsulta(rst2)

								End If

								sActcon = "Consultar.asp?Codigo=" & codobjetoL & "&Tipo=" & wtipoLc
								sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
								sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3



								sHref    = "&nbsp;"

								sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;" & Chr(34)


								slnktit  = wNomL
								wNomL    = pStyle & "&nbsp;<u><i>" & wNomL & "</i></u>&nbsp;" & pStyleF
								sImagen  = wNomL



								sLink    = OFV.GenLink("", "nt title='" & slnktit & "' ",GenKeyLnks(sHref), "", "", sFuncion, "NNNNNN", "si")
								descri   = sLink
								Call ofv.cerrarconsulta(rst2)

							Else

								If  Not rst.eof Then

									wUrl  = rst("Lurl") & ""

									If Not isnull(wUrl) Then
										sHrefu = instr(wUrl,"://")
										If sHrefu > 0 Then
											sHrefu = wUrl
										Else
											sHrefu = "HTTP://" & wUrl
										End If
									Else
										sHrefu = "javascript:void(0);"
									End If

									slnktit  = rst("LurlDesc")
									wNomL    = pStyle & "&nbsp;<u><i>" & rst("LurlDesc") & "</i></u>&nbsp;" & pStyleF
									sImagen  = wNomL

									descri = ofv.GenLink("","nt title='" & slnktit & "' ",GenKeyLnks(sHrefu),"_blank","","","NNNNNN","si")

								End If

							End If


						End If
						Call ofv.cerrarconsulta(rst)

					End If

				End If



				If wasotxt = "S" Then
					' response.write wasotxtstr & "  " & wasotxtcan

					If len(sTextoV) > 0 Then

						If  wasotxtstr > 0  And wasotxtcan > 0 Then

							sTextoV = left(sTextoV,wasotxtstr - 1) & "" & replace(descri,"NNNNNN",pStyle & "&nbsp;<u><b style='color:" & Session("APLFT") & "' >" & mid(sTextoV,wasotxtstr,wasotxtcan) & "</b></u>&nbsp;" & pStyleF) & "" & right(sTextoV,len(sTextoV) - (wasotxtstr + wasotxtcan - 1))

						End If

						sTextoV =  replace(sTextoV,chr(13)," <br>")

					End If


				End If


				If wasotxt = "N" Then

					If len(sTextoV) > 0 Then

						'sTextoV = supersec & "---" & ta& replace(sTextoV,chr(13)," <br>")
						sTextoV =  replace(sTextoV,chr(13)," <br>")

					End If

					If  sImagen <> "NO" Then sTextoV = sTextoV & replace(descri,"NNNNNN",sImagen)

				End If


				If session("APLMDCLI")="BC" Then
					Celda   = Celda & ofv.GenCelda("",bspan,Ptalig & " valign=TOP  ","","",pbgcol,"",pStyle & PsangS  & sTextoV & pStyleF,"si")
				Else
					Celda   = Celda & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;' ","","",pbgcol,"",pStyle & PsangS  & sTextoV & pStyleF,"si")
				End If

				sTextoVKK = ""

				If Marca = "CA" Then

					If  RCHEK = "S" Then
						Filatt  = ""
						Celdakk = ""
						Celdakk = Celdakk & ofv.GenCelda("","utg colspan=12 style='text-align:justify;color:#0054a9;' ","","30%","","","","<u>Version Anterior:</u>","si")
						Filatt  = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

						sTextoVKK = rs3(8)
						Celdakk = ""
						Celdakk = Celdakkp
						''             CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"",PsangS  & sTextoVKK,"si")
						CeldaKK = CeldaKK & ofv.GenCelda("",replace(bspan,"aaa","utg"),Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"",PsangS  & sTextoVKK,"si")
						Filatt  = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

					End If

				End If

				If PnumS = "" Then
					'        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;x1","si")
				End If

				If PVin = "" Then
					'      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;x2","si")
				End If

			End If

			' IMAGENES            --------------------------------------------------

			If stip = "H" Then


				sreduc = ""
				sbord = ""
				salng = ""

				If  ars(12,inx) <> "" And ars(12,inx) <> "100" Then sreduc = " width=" & ars(12,inx) & "% "
				If  ars(15,inx) <> "" Then
					If ars(15,inx) = "SI" Then sbord = " border=1 "
				End If

				If  ars(16,inx) <> "" Then
					If ars(16,inx) = "CENTRADO" Then salng = " center"
					If ars(16,inx) = "DERECHA" Then salng = " right"
					If ars(16,inx) = "IZQUIERDA" Then salng = " left"
				End If

				If ars(9,inx) <> "" Then
					If session("EMPMUN") = "S" Then
						slink = "<img src=" & chr(34) & GenKeyLnks("Emp" & SESSION("D" & FNCODOBJ & "-0" ) & "/estructura/" & ars(9,inx))  & chr(34) & " "
					Else
						slink = "<img src=" & chr(34) & GenKeyLnks( "estructura/" & ars(9,inx))  & chr(34) & " "
					End If

					slink = slink & sreduc & sbord & " >"
					slink2 = slink
					If session("EMPMUN") = "S" Then
						slink = "<a href='" & GenKeyLnks("Emp" & SESSION("D" & FNCODOBJ & "-0" ) & "/estructura/" & ars(9,inx))  &  "' target='_blank' >" & slink2 & "</a>"
					Else
						slink = "<a href='" & GenKeyLnks("estructura/" & ars(9,inx))  &  "' target='_blank' >" & slink2 & "</a>"
					End If


				Else
					slink = ""
				End If

				Celda = Celda & ofv.GenCelda("",bspan,salng,"","","","",slink,"si")
				Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")

				slinkkk = ""


				If Marca = "CA" Then

					If   RCHEK = "S" Then
						If rs3(9) & "" <> "" Then
							Filatt  = ""
							CeldaKK   = ""
							CeldaKK   = CeldaKK & ofv.GenCelda("","utg colspan=12 "," valign=TOP style='text-align:justify;color:#0054a9;' ","","","","", "<u>Version Anterior:</u>","si")
							Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

							CeldaKK   = ""
							If session("EMPMUN") = "S" Then
								slinkkk = "<img src=" & chr(34) & GenKeyLnks("Emp" & SESSION("D" & FNCODOBJ & "-0" ) & "/estructura/" & rs3(9))  & chr(34) & " "
							Else
								slinkkk = "<img src=" & chr(34) & GenKeyLnks( "estructura/" & rs3(9))  & chr(34) & " "
							End If

							slinkkk = slinkkk & sreduc & sbord & " >"
							slink2kkk = slinkkk

							CeldaKK   = CeldaKK & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"", slinkkk & "","si")
							''               CeldaKK   = CeldaKK & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"", slinkkk & "","si")
							Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

						End If

					End If

				End If

			End If

			' Tablas           --------------------------------------------------

			If stip = "I" Then

				slink = PasteTabla(ars(10,inx),nCodigo)
				Celda = Celda & ofv.GenCelda("",bspan,"","","","","",slink,"si")
				Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")

			End If

			' COMPLEMENTOS            --------------------------------------------------

			If stip = "J" Then

				'Descripcion
				sreduc = ""
				sbord = ""
				salng = ""
				If ars(12,inx) <> "" Then sreduc = " width=" & ars(12,inx) & "% "
				If ars(15,inx) <> "" Then
					If ars(15,inx) = "SI" Then sbord = " border=1 "
				End If
				If ars(16,inx) <> "" Then
					If ars(16,inx) = "CENTRADO" Then salng = " center"
					If ars(16,inx) = "DERECHA" Then salng = " right"
					If ars(16,inx) = "IZQUIERDA" Then salng = " left"
				End If

				If (session("APLMDCLI") = "BGAL" Or session("CORMODUFORM")="S") And ars(32,inx & "")  <> "" Then

					strsql = "SELECT *  From objlnk where objid = " & ars(32,inx & "")
					'  response.write strsql
					Set rst = ofv.crearconsultaEx(StrSql,cn,1,volver)
					If Not rst.eof Then

						wasotxt     = rst(14) & ""
						wasotxtstr  = CLNG("0" & rst(15))
						wasotxtcan  = CLNG("0" & rst(16))
						wasotxtdat  = rst(17) & ""

						If wasotxt = "" Then wasotxt = "N"



						wNV1    = rst(11) & ""
						wNV2    = rst(12) & ""
						wNV3    = rst(13) & ""

						If  WNV1 = "" Then WNV1 = "0"
						If  WNV2 = "" Then WNV2 = "0"
						If  WNV3 = "" Then WNV3 = "0"

						codobjetoL = clng(rst("codobjetoL"))
						wtipoLc    = rst(7)
						If trim(rst("Ltipo")) = "KF" Then

							wUrlDesc = rst(9) & ""


							wNomL    = wUrlDesc
							If wNomL = "" Then
								wNomL = "S/D"


								Strsql = "Select * From csGenObjetos where estado <> 4 and codobjeto = " & codobjetoL
								'      RESPONSE.WRITE STRSQL
								Set rst2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								If Not rst2.eof Then
									wNomL = rst2(1)
								Else
									wNomL =""' "S/D"
								End If
								Call ofv.cerrarconsulta(rst2)

							End If

							sActcon = "Consultar.asp?Codigo=" & codobjetoL & "&Tipo=" & wtipoLc
							sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
							sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3



							sHref    = "&nbsp;"
							' response.write sActcon
							sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;" & Chr(34)


							slnktit  = wNomL
							wNomL    = pStyle & "&nbsp;<u><i>" & wNomL & "</i></u>&nbsp;" & pStyleF
							sImagen  = wNomL

						End If
					End If
					'sImage = "<font face='corflaticon' size=5 color=#f7aa22 >N</font>"
					sImage = "<i class='las la-eye' ></i>"
					sLink    = OFV.GenLink("", "docs title='" & slnktit & "' ",GenKeyLnks(sHref), "", "", sFuncion,sImage & wNomL, "si")


				ElseIf ars(9,inx) <> ""  Then




					Strsql = "Select codobjetol,tipobjetol,* From objlnk Where ltipo = 'KL' and objid = 0" & ars(9,inx)
					' response.write strsql
					Set rsCompoM = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If  Not rsCompoM.eof Then

						wtipoL   = rsCompoM(8)
						wtipoLc  = rsCompoM(9) & ""

						wNV1    = rsCompoM(13) & ""
						wNV2    = rsCompoM(14) & ""
						wNV3    = rsCompoM(15) & ""

						If  WNV1 = "" Then WNV1 = "0"
						If  WNV2 = "" Then WNV2 = "0"
						If  WNV3 = "" Then WNV3 = "0"



						Strsql = "Select nombre,TIPOCOMP From csacomple Where codcomp = " & rsCompoM(0)
						' response.write strsql & chr(13)
						Set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rsComponentes.eof Then
							sComponente = rsComponentes(0)
							sTIPOCMP    = rsComponentes(1)

						End If
						Call ofv.cerrarconsulta(rsComponentes)


						sInput = sComponente

						If len(sInput) > 50 Then
							sInput = left(sInput,47) & "..."
						End If

						sARCHIVO = ""
						sARCHmas = "N"
						'            Strsql = "SELECT   * FROM COMPIMAG where codcomp = 0" & rsCompoM(0) & " ORDER BY SECUENCIA, OBJID "

						Nv=WNV3
						If nv=0 Then
							Nv=wnv2
							If nv=0 Then
								nv=wnv1
							End If
						End If

						Strsql = "SELECT   * FROM COMPIMAG where objid = 0" & NV & " ORDER BY SECUENCIA, OBJID "

						' response.write strsql
						Set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						Do until rsComponentes.eof
							sARCHIVO = rsComponentes(3)
							rsComponentes.movenext
							If Not rsComponentes.eof Then
								sARCHmas = "S"
								Exit Do
							End If
						Loop
						Call ofv.cerrarconsulta(rsComponentes)


						sInput = "&nbsp;&nbsp;<B>" & replace(sInput," ","&nbsp;") & "&nbsp;" & "</b>"

						If sARCHmas = "S" Then

							sInput = sInput & "&nbsp;&nbsp;(mas imagenes)"
						End If

						'sImage = "<font face='corflaticon' size=5 color=#f7aa22 >N</font>" & sInput
						sImage = "<i class='las la-eye' ></i>" & sInput
						sHref       = "EditarImagen2.asp?ID=" & wtipoL
						sHref       = sHref & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

						sHref = GenKeyLnks(sHref)
						' sHref = GenKeyLnks("imagencont4.asp?ID=" & rsCompoM(0))
						sLink2 = ofv.GenLink("","docs title='Visualizar el componente' ",sHref,"_new","","",sImage,"si")

						slink = "<BR>" & sLink2

						If SARCHIVO <> "" Then
							slink = slink & "<BR><BR><img src=" & chr(34) & GenKeyLnks("COMPONENTES/" & SARCHIVO)  & chr(34) & " "
							slink = slink & sreduc & sbord & " >"
						End If

						slink = slink & "<BR>"



					Else
						slink = ""

					End If
					Call ofv.cerrarconsulta(rsCompoM)







				Else
					slink = ""
				End If

				Celda = Celda & ofv.GenCelda("",bspan,salng,"","","","",slink,"si")
				Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")

				slinkxx = ""

				If Marca = "CA" Then
					'sImage = "<font face='corflaticon' size=5 color=#f7aa22 >N</font>" & sInput
					sImage = "<i class='las la-eye' ></i>" & sInput
					sHref       = "EditarImagen2.asp?ID=" & wtipoL
					sHref       = sHref & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

					sHref = GenKeyLnks(sHref)
					sLink2xx = ofv.GenLink("","docs title='Visualizar el componente' ",sHref,"_new","","",sImage,"si")

					slinkxx = "<BR>" & sLink2xx

					If SARCHIVO <> "" Then
						slinkxx = slinkxx & "<BR><BR><img src=" & chr(34) & GenKeyLnks("COMPONENTES/" & SARCHIVO)  & chr(34) & " "
						slinkxx = slinkxx & sreduc & sbord & " >"
					End If

					slinkxx = slinkxx & "<BR>"
					If   RCHEK = "S" Then
						Filatt  = ""
						Celdakk = ""
						Celdakk = Celdakk & ofv.GenCelda("","utg colspan=3 style='text-align:justify;color:#0054a9;' ","","30%","","","","<u>Version Anterior:</u>","si")
						Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")
						Celdakk = ""
						''                CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & "dt valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"", slinkxx & "","si")
						CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & "dt valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"", slinkxx & "","si")
						Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

					End If

				End If

			End If

			'------------------------------------------------


			Select Case stip
				Case "A"
				sant     = "A"
				sVinoAnt = sant

				Case "B"
				sant     = "B"
				sVinoAnt = sant
				Case "C"
				sant     = "C"
				sVinoAnt = sant
				Case "D"
				sant     = "D"
				sVinoAnt = sant
				Case "E"
				sant     = "E"
				sVinoAnt = sant
				Case "F"
				sant     = "F"
				sVinoAnt = sant
				Case Else

				sVinoAnt = stip

			End Select


			Fila = ofv.GenRow("","","","","","","",celda,"si")


			wcantlin = wcantlin + 1

			If   wcantlin > 20 Then
				wcantlin = 0
			End If



			If Mostrar <> "NO"  Then

				FNSHTML = FNSHTML &  Fila



				''response.write marca & " 03"

				If PEspa = "" Then
					PEspa      = 0
					PEspHeight = ""
				ElseIf PEspa = 0 Then
					PEspa      = 0
					PEspHeight = ""
				Else
					PEspa      = 0 + PEspa
					PEspHeight = 9 * PEspa
				End If



				celda = ""
				For i=1 To 12
					Celda = Celda & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
				Next
				Fila = ofv.GenRow("","","","",PEspHeight,"","",celda,"si")
				FNSHTML = FNSHTML &  Fila


				If sVinoAnt > "F" And stip < "G" Then
					celda = ""
					For i=1 To 10
						Celda = Celda & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
					Next
					Fila  = ofv.GenRow("","","","","12","","",celda,"si")
					FNSHTML = FNSHTML &  Fila

				End If

				If Marca = "CA" Then
					If  RCHEK = "S" Then
						''            FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Celdakk,"si")
						celda = ""
						For i=1 To 12
							Celda = Celda & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
						Next
						Filatt = Filatt & ofv.GenRow("","","","","1","","",celda,"si")

						SArmatt = ofv.GenTabla("","AAA style='BACKGROUND-COLOR:#f9f9f9;border-style:solid;border-width:1;border-color:#0054a9;'  ","","80%","","0","0","2","2","","",Filatt,"si")

						Celdakk = ""
						Celdakk = Celdakk & ofv.GenCelda("","ut colspan=12 align=CENTER ","","","","","",SArmatt,"si")
						''  Fila = Fila & ofv.GenRow("","","","","","","",Celdakk,"si")


						FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Celdakk,"si")
						''  FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Fila,"si")

						Filatt = ""

					End If
				End If


			End If

			''    FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Fila,"si")


		Next

		If inx > UBound(ars,2) Then
			Exit Do

		End If

		ssale = "no"
		Do until ssale = "si"

			If     sant = "A" Then

				sant = "X"
			ElseIf sant = "B" Then

				sant = "A"
			ElseIf sant = "C" Then

				sant = "B"
			ElseIf sant = "D" Then

				sant = "C"
			ElseIf sant = "E" Then

				sant = "D"
			ElseIf sant = "F" Then

				sant = "E"
			End If

			If sant = "X" Then
				ssale = "si"
			End If

		Loop

	Loop

	celda = ""
	For i=1 To 12
		wwidth = "1%"
		If i = 9 Then wwidth = "90%"
		Celda = Celda & ofv.GenCelda("","ut","",wwidth,"","","","&nbsp;","si")
	Next

	Fila = ofv.GenRow("","","","","","","",celda,"si")
	FNSHTML = FNSHTML &  Fila




	' call ofv.cerrarconsulta(rs)


	celda = ""




	FNSHTML = FNSHTML & ofv.GenFTabla()


	''
	GENFMTESTRUCTURA = "<div id='zglos' class='norma-card' style='border-radius: 0px 0px 10px 10px;' >" & FNSHTML & "</div>"


End Function



'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
Function ConvNumer(CNumero,CTipo)

	CLetra = "abcdefghijklmnopqrstuvwxyz"

	Dim CRuni(9)
	CRuni(1) = "I"
	CRuni(2) = "II"
	CRuni(3) = "III"
	CRuni(4) = "IV"
	CRuni(5) = "V"
	CRuni(6) = "VI"
	CRuni(7) = "VII"
	CRuni(8) = "VIII"
	CRuni(9) = "IX"

	Dim CRdec(9)
	CRdec(1) = "X"
	CRdec(2) = "XX"
	CRdec(3) = "XXX"
	CRdec(4) = "XL"
	CRdec(5) = "L"
	CRdec(6) = "LX"
	CRdec(7) = "LXX"
	CRdec(8) = "LXXX"
	CRdec(9) = "XC"

	Dim CRcen(9)
	CRcen(1) = "C"
	CRcen(2) = "CC"
	CRcen(3) = "CCC"
	CRcen(4) = "CD"
	CRcen(5) = "D"
	CRcen(6) = "DC"
	CRcen(7) = "DCC"
	CRcen(8) = "DCCC"
	CRcen(9) = "CM"


	Select Case Ctipo
		Case "a","A"
		Ci = 0
		CR = CNumero

		For ci=0 To 28
			If CR < 28 Then Exit For
			CR = CR - 27
		Next

		If ci > 27 Then
			ConvNumer = "XX"
		Else
			If ci > 0 Then ConvNumer = mid(CLetra,ci,1)
			ConvNumer = ConvNumer & mid(CLetra,CR,1)
		End If
		If Ctipo = "A" Then ConvNumer = Ucase(ConvNumer)

		Case "I","i"
		CR = CNumero
		Ci = len(CR)
		If Ci < 1 Or ci > 3 Then
			ConvNumer = "AAAA"
		Else
			cx = 1
			For CL=ci To 1 Step -1
				Select Case cx
					Case 1
					ConvNumer = CRuni(mid(cr,cl,1))
					Case 2
					ConvNumer = CRdec(mid(cr,cl,1)) & ConvNumer
					Case 3
					ConvNumer = CRcen(mid(cr,cl,1)) & ConvNumer
				End Select

				cx = cx + 1
			Next
			If Ctipo = "i" Then ConvNumer = Lcase(ConvNumer)
		End If
	End Select


End Function


Function reemplazaspan9(textoz)

	jsspan=""

	Dim Position
	Dim WordLength
	Word="<SPAN id="
	Position = InStr(1, Textoz, Word,0)

	WordLength = Len(Word)
	Do While Position


		idd= InStr(Position + WordLength, Textoz," ",0)
		'idd= InStr(Position + WordLength, Textoz," ",0)
		total=idd- ( position+ + WordLength )
		idPerfil =mid(textoz,Position + WordLength,total)
		sql="select a.descri from Tipoperfil a inner JOIN ecoperfil b on( b.codperfil=a.tipoPERfil and b.objid=0" & idPerfil & ")"

		'       RESPONSE.WRITE SQL
		Set xpf=ofv.crearconsultaEx(sql,cn,1,parametros)
		If Not xpf.eof Then
			sfin=InStr(idd, Textoz,"</SPAN>",0)
			xini=InStr(idd, Textoz,"xspanx>",0)
			'         TEXTOLD=MID(TEXTOZ,xini+ len("xspanx>"),SFIN-xini-1)
			'        response.write textold & "222"
			'   Textoz=REPLACE(Textoz,TEXTOLD,xpf(0))
			jsspan = jsspan & " document.getElementById('"&idPerfil &"').innerHTML='"&xpf(0)&"';" & chr(13)
		End If
		Position = InStr(Position + WordLength, Textoz,Word,0)
	Loop
	jsspan =jsspan & ""
	reemplazaspan = textoz
End Function

%>