<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
Function GeneraStaff()

	fila = ""

	bordert = ""
	borderl = ""
	borderr = ""
	bstilo = ""
	bcolor = ""
	ccolor = "#ededed"
	cstilo = "border-style:solid;border-width:1;border-Color:#000000;"
	cbordert = "border-top-style:solid;border-top-width:1;border-top-Color:#000000;"
	cborderl = "border-left-style:solid;border-left-width:1;border-left-Color:#000000;"
	cborderr = "border-right-style:solid;border-right-width:1;border-right-Color:#000000;"


	For xinS = 1 To UBound(vecsecS) Step 2
		borderl = cborderl
		bordert = cbordert
		bstilo = cstilo
		celda = OFV.gencelda("", "aaa colspan=9 ", "", "", "", "", "", "&nbsp;", "si")
		celda = celda & OFV.gencelda("", "aaa colspan=9  style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
		fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

		celda = OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
		shref = "organigramaUNGraf.asp?ID=" & vecsecS(xinS) & "&Sector=" & OFV.ConvertirCar(vecsecSD(xinS), " ", "_")
		sinput = "<font face=verdana color=#6699cc size=1>" & vecsecSD(xinS) & "</font>"
		slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

		celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", ccolor, "", slink, "si")
		celda = celda & OFV.gencelda("", "aaa colspan=5 ", "", "", "", "", "", "&nbsp;", "si")
		celda = celda & OFV.gencelda("", "aaa colspan=5 style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
		If xinS < UBound(vecsecS) Then
			shref = "organigramaUNGraf.asp?ID=" & vecsecS(xinS + 1) & "&Sector=" & OFV.ConvertirCar(vecsecSD(xinS + 1), " ", "_")
			sinput = "<font face=verdana color=#6699cc size=1>" & vecsecSD(xinS + 1) & "</font>"
			slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

			celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", ccolor, "", slink, "si")
		Else
			celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 ", "center", "18%", "", "", "", "&nbsp;", "si")
		End If
		celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
		fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

		bordert = cbordert
		borderl = cborderl
		celda = OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
		celda = celda & OFV.gencelda("", "aaa colspan=5  style='" & bordert & "'", "", "", "", "", "", "&nbsp;", "si")
		bordert = ""
		borderl = ""
		If xinS < UBound(vecsecS) Then bordert = cbordert
		If (xinS + 1) < UBound(vecsecS) Or UBound(vecsec) > 0 Then borderl = cborderl
		celda = celda & OFV.gencelda("", "aaa colspan=5  style='" & bordert & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
		celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
		fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

	Next

	GeneraStaff = fila

End Function

Function GeneraResto()

	fila = ""

	ccolor = "#ffffff"
	cstilo = "border-style:solid;border-width:1;border-Color:#000000;"
	cbordert = "border-top-style:solid;border-top-width:1;border-top-Color:#000000;"
	cborderl = "border-left-style:solid;border-left-width:1;border-left-Color:#000000;"
	cborderr = "border-right-style:solid;border-right-width:1;border-right-Color:#000000;"

	xin = 1

	vniv = vecsecn(xin)

	Do Until xin > UBound(vecsec)

		celda = OFV.gencelda("", "aaa colspan=9 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ", "", "", "", "", "", "&nbsp;", "si")
		celda = celda & OFV.gencelda("", "aaa colspan=9  style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;" & cborderl & "'", "", "", "", "", "", "&nbsp;", "si")
		fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

		vnivd = 0

		xnivT = "F"

		For xinn = xin To UBound(vecsecn)

			If vniv <> vecsecn(xinn) Then
				vniv = vecsecn(xinn)
				xnivT = "N"
				Exit For
			End If
			vnivd = vnivd + 1
		Next



		'''response.write "xin=" & xin & " cant en nivel=" & vnivd & "<br>"
		xinn = xin

		xpass = 0

		For xin = xinn To (xinn + vnivd - 1) Step 4
			'response.write  vecsecd(xin) & "<br>"
			xinnp = 4 * xpass

			xpass = xpass + 1


			bordert = ""
			borderl = ""
			borderr = ""
			bstilo = ""
			bcolor = ""


			borderl = cborderl
			celda = OFV.gencelda("", "aaa colspan=9 ", "", "", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "aaa colspan=9  style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
			fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
			borderl = ""

			vcant = vnivd - xinnp
			vcantr = 5

			If xnivT = "F" Then
				vcantr = vcant
			End If

			celda = OFV.gencelda("", "aaa colspan=3 ", "", "", "", "", "", "&nbsp;", "si")

			If vcantr > 2 Then
				bordert = cbordert
				borderl = cborderl
				borderr = ""
			End If

			celda = celda & OFV.gencelda("", "aaa colspan=4  style='" & bordert & borderl & borderr & "' ", "", "", "", "", "", "&nbsp;", "si")

			If vcantr > 1 Then
				bordert = cbordert
				borderr = ""
				borderl = ""
				If (vcantr = vcant) And (vcant = 2 Or vcant > 3) Then
					borderl = cborderl
				End If

				If (vcantr <> vcant) And (vcant > 2) Then
					borderl = cborderl
				End If

			End If

			celda = celda & OFV.gencelda("", "aaa colspan=2  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")

			borderr = ""
			If vcantr = 1 Or vcantr = 3 Or vcantr > 4 Then
				borderl = cborderl
			Else
				borderl = ""
			End If

			If vcant > 1 Then
				bordert = cbordert
			Else
				bordert = ""
			End If

			celda = celda & OFV.gencelda("", "aaa colspan=2  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")

			bordert = ""
			borderl = ""
			borderr = ""
			If (vcantr = vcant) And (vcant = 2 Or vcant > 3) Or (vcantr <> vcant) And (vcant > 3 ) Then
				''''''   If vcant = 2 Or vcant > 3 Then
				borderl = cborderl
			End If

			If ((vcantr = vcant) And (vcant > 2)) Or ((vcantr <> vcant) And (vcant > 1)) Then
				bordert = cbordert
				borderr = cborderr
			End If

			celda = celda & OFV.gencelda("", "aaa colspan=4  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "aaa colspan=3  ", "", "", "", "", "", "&nbsp;", "si")
			fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

			bordert = ""
			borderl = ""
			borderr = ""
			bstilo = ""
			bcolor = ""
			slink = "&nbsp;"

			celda = OFV.gencelda("", "vt colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

			If ((vcant = vcantr) And (vcant > 2)) Or ((vcant <> vcantr)) Then
				bstilo = cstilo
				bcolor = ccolor
				shref = "organigramaUNGraf.asp?ID=" & vecsec(xin) & "&Sector=" & OFV.ConvertirCar(vecsecd(xin), " ", "_")
				sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xin) & "</font>"
				slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
			End If

			celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2  style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
			celda = celda & OFV.gencelda("", "vt colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

			If ((vcant = vcantr) And (vcant = 2 Or vcant > 3)) Or ((vcant <> vcantr) And (vcant > 2)) Then
				bstilo = cstilo
				bcolor = ccolor
				If vcant = 2 Then
					xinW = xin
				Else
					xinW = xin + 1
				End If



				shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
				sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
				slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
			Else
				bstilo = ""
				bcolor = ""
				slink = "&nbsp;"
			End If

			celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")

			If ((vcant = vcantr) And (vcant = 3 Or vcant = 1))  Then

				'''''''''' If vcant = 3 Or vcant = 1 Then
				bstilo = cstilo
				bcolor = ccolor
				If vcant = 3 Then
					xinW = xin + 1
				Else
					xinW = xin
				End If
				shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
				sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
				slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
				celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")

			Else
				If vcantr > 4 Then borderl = cborderl
				celda = celda & OFV.gencelda("", "vt  rowspan=2 ", "", "", "", "", "", "&nbsp;", "si")
				celda = celda & OFV.gencelda("", "vt  rowspan=2 style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
			End If


			If ((vcant = vcantr) And (vcant = 2 Or vcant > 3)) Or ((vcant <> vcantr) And (vcant > 3)) Then

				''''''  If vcant = 2 Or vcant > 3 Then
				bstilo = cstilo
				bcolor = ccolor
				If vcant > 2 Then
					xinW = xin + 2
				Else
					xinW = xin + 1
				End If

				shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
				sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
				slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
			Else
				bstilo = ""
				bcolor = ""
				slink = "&nbsp;"
			End If

			celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
			celda = celda & OFV.gencelda("", "vt colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

			If ((vcant = vcantr) And (vcant > 2)) Or ((vcant <> vcantr) And (vcant > 1)) Then

				'''''''   If vcant > 2 Then
				bstilo = cstilo
				bcolor = ccolor
				If vcant > 3 Then
					xinW = xin + 3
				ElseIf vcant > 2 Then
					xinW = xin + 2
				Else
					xinW = xin + 1
				End If
				shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
				sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
				slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
			Else
				bstilo = ""
				bcolor = ""
				slink = "&nbsp;"
			End If

			celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
			celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
			fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

			celda = OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
			fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")


		Next

		If xnivT = "Ff" Then

			celda = OFV.gencelda("", "aaa colspan=9 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ", "", "", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "aaa colspan=9  style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;" & cborderl & "'", "", "", "", "", "", "&nbsp;", "si")
			fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

		End If

		xin = xinn + vnivd

	Loop


	GeneraResto = fila


End Function


ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else



	Generar = ""

	'response.buffer=false
	Set cn = OFV.conectar(OFV.strconn2)
	Set cn2 = OFV.conectar(OFV.strconn0)
	Set cn3 = OFV.conectar(OFV.strconn0)



	If request.querystring("mseg")<>"" Then
		session("mseg")=request.querystring("mseg")
	End If

	nID = ""

	sSector = ""
	nSeg = ""

	nID = Request.querystring("ID")




	If SESSION("MSEG") = "E" Then

		'  sSQL = "SELECT sector,nivel FROM usuarios where id = '" & session("usuario") & "'"
		' Set rs = OFV.crearconsulta2(sSQL, cn2)
		'     If Not rs.EOF Then
		nidu  = session("SGU_idcostcenter") 'rs(0)
		nivel = session("SGU_idnivel") 'rs(1)
		'    end if
		'   Call OFV.cerrarconsulta(rs)

		'xxx=getinfouser(session("usuario"))

		If nID = "" Then
			nID = nidu
		End If


	Else
		nivel=0

		sSector = Request.querystring("Sector")
		sSector = OFV.ConvertirCar(sSector, "_", " ")
		sCodin = "&nbsp;"
		If nID = "" And Request.Form.Count > 0 Then
			nID = Request.Form(1)

		End If
	End If

	nid = "0" & Replace(nid,"nsnc","")


	If Not nid > "0" Then

		''   sHTML = OFV.ReportHeader("<FONT SIZE=2 COLOR=#6699CC><B>" & Session("Cliente") & "</B></FONT>") & ofv.Repbody("CORDPO - ORGANIGRAMA")

		sHTML = OFV.MenuHeader("","")

		shtml=shtml & "<br><br><br><br><br><center><p><b><font face=verdana size=4 color=#993333>DEBE SELECCIONAR UN AREA O SECTOR<br>PARA GRAFICAR.</FONT></P>" & Chr(13)
		sHTML = sHTML  & "<br></body></html>"
	Else


		sSECSITE  = 1
		sSECNIVEL = 0
		SNIVELA   = "S/D"

		If nID <> "" Then
			sSQL = "Select descri,codinter,ECONIV From eco Where eco.eco = " & nID
			Set rs = OFV.crearconsultaEx(sSQL, cn, 1, "")
			If Not rs.EOF Then
				sSector = rs(0)
				sCodin = rs(1)
				sSECSITE  = Request.Form("SECSITE")
				sSECNIVEL = Request.Form("SECNIVEL")


				If  sSECSITE <> "" Then
					SESSION("SSITE") = sSECSITE
				Else
					sSECSITE = SESSION("SSITE")
				End If

				If sSECNIVEL  = "" Then sSECNIVEL  = 0


				sSQL = "SELECT * FROM ECOSITE where ECO = 0" & nID & " AND ECOSITE = 0" & sSECSITE
				Set rs1 = OFV.crearconsulta2(sSQL, cn)
				If  rs1.EOF Then
					sSQL = "SELECT * FROM ECOSITE where ECO = 0" & nID
					Set rs1 = OFV.crearconsulta2(sSQL, cn)
					If  Not rs1.EOF Then
						sSECSITE = RS1(2)
					End If
				End If

				sSUCA = ""
				sSUCN = ""

				sSQL = "SELECT * FROM SITES where OBJID = 0" & sSECSITE
				Set rs1 = OFV.crearconsulta2(sSQL, cn)
				If  Not rs1.EOF Then
					sSUCA = RS1(1)
					sSUCN = RS1(3) & "&nbsp;"
				End If



				sSQL = "SELECT * FROM NIVELJERAR where CODNIVEL = 0" & RS(2)
				Set rs1 = OFV.crearconsulta2(sSQL, cn)
				If Not rs1.EOF Then
					SNIVELA      = RS1(1)
					'''  response.write "NIV " & SNIVELA
				End If

			End If
			Call OFV.cerrarconsulta(rs)
		End If


		' response.write NID & "----" & sSector
		sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=2>&#0128;</FONT></B>"
		sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"
		sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
		If nID <> "" Then
			''  sHTML = OFV.ReportHeader("<FONT SIZE=2 COLOR=#6699CC><B>" & Session("Cliente") & "</B></FONT>") & ofv.Repbody("CORDPO - ORGANIGRAMA")

			sHTML = OFV.MenuHeader("","")
		    shtml = replace(shtml,"<body ","<body onload=largo(); ")
			shtml=shtml &  Chr(13) & "<script type=" & Chr(34) & "text/javascript"& Chr(34) & " src='IncFile/dragdrop.js' ></script>" & Chr(13)

			shtml=shtml & "<script language=javascript>" & Chr(13)
			shtml=shtml & "function mostrarmenu(xusr,XUSRN){" & Chr(13)

			shtml=shtml & "}" & Chr(13)
			shtml=shtml & "</script>" & Chr(13)
			shtml = shtml & Chr(13) & "<script type='text/javascript' src='incfile/ECOTreeN.js' ></script><link type='text/css' rel='stylesheet' href='incfile/ECOTree.css' />"
			shtml=  shtml & "<script type='text/javascript' src='JSGRAF/jspdf.js'></script>"


			shtml=  shtml & "<script type='text/javascript' src='JSGRAF/jspdf.plugin.addimage.js'></script> "
			shtml=  shtml & " <script type='text/javascript' src='JSGRAF/jspdf.source.js'></script>"

			sinput = "<input type=button value='Subir' style='color:#000000;font-weight:bold;background-color:#ccddee;' onclick='document.OPSUBE.submit();'  ><form id='OPSUBE' name='OPSUBE' action='" & shref & "' method='post' style='display:none;' >"
			slink = ofv.GenLink("", "nt", shref, "", "", "", sinput, "si")





			If SESSION("MSEG")="E" Then
				visible =" display:none;"
				sImagen4si="&nbsp;"
			Else
				visible=""
			End If






			xin = 0
			ReDim Preserve vecsec(xin)
			ReDim Preserve vecsecd(xin)
			ReDim Preserve vecsecn(xin)

			'Sector Principal



			sSQL = "SELECT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ER.CODREL FROM ECO E , ECOREL ER, ECO  ECO_1 where E.eco = ER.ECODEPEND"
			sSQL = sSQL & " and ER.ECOMASTER = ECO_1.eco and E.eco = " & nID
			'   response.write strsql
			Set rs = OFV.crearconsulta2(sSQL, cn)
			If rs.EOF Then
				vecsec(xin)  = 0
				vecsecd(xin) = sSector
				vecsecn(xin) = 0

				 wwwsettree=""
			     wwwsettree = wwwsettree & "myTree.add2(0,-1,'"&sSector&"',120,60,'" & session("CMFMENU") &"','#000000','" & shref & "',null,'1','NO');"
			Else
				'''''''''''''''
				'sSECSITE
                wwwsettree=""
                shref = "OrganigramaUNgraf.asp?id=" & rs(0)
			    wwwsettree = wwwsettree & "myTree.add2(0,-1,'" & sSector & "',120,60,'" & session("CMFMENU") &"','#000000','" & shref & "',null,'1','NO');"

				Do Until RS.EOF

					'''''''''''''''
					SCODREL = RS(9)

					sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL
					'          response.write ssql
					Set rs2 = OFV.crearconsulta2(sSQL, cn)
					If Not rs2.EOF Then
						sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL & " AND SITE = " & sSECSITE
						'  response.write strsql
						Set rs2 = OFV.crearconsulta2(sSQL, cn)
						If Not rs2.EOF Then
							vecsec(xin)  = CLng(rs(0))
							vecsecd(xin) = sSector
							vecsecn(xin) = rs(5)
							Exit Do
						End If
					Else
						If sSECSITE = 1 Then
							vecsec(xin)  = CLng(rs(0))
							vecsecd(xin) = sSector
							vecsecn(xin) = rs(5)
							Exit Do
						End If
					End If


					''''''''''''
					RS.MOVENEXT
				Loop
				If rs.EOF Then
					vecsec(xin)  = 0
					vecsecd(xin) = sSector
					vecsecn(xin) = 0
				End If

				'''''''''''''

			End If




			sSQL = "SELECT DISTINCT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ER.CODREL,ER.ECOMASTER,ER.ECODEPEND "
			sSQL = sSQL & " FROM ECO E, ECOREL ER, ECO  ECO_1, segmento, ECOSITE ES "
			sSQL = sSQL & " where E.eco = ER.ECOMASTER and "
			sSQL = sSQL & " ER.ECODEPEND = ECO_1.eco and "
			sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and ES.ECO = ECO_1.ECO AND "
			sSQL = sSQL & " E.eco = " & nID & " AND (ES.ECOSITE = 0" & SESSION("SSITE")
			sSQL = sSQL & " OR ES.ECOSITE = 0" & sSECSITE & ") and "
			sSQL = sSQL & " segmento.descri <> 'STAFF' order by ECO_1.ECONIV"

			'  response.write ssql
			Set rs = OFV.crearconsulta2(sSQL, cn)

			'Texto de Sectores Continuadores y Lineas
			pad = 1
			Do Until rs.EOF
				''''''''''''''''
				SCODREL = RS(9)
				'  smaster = RS(10)
				'  sdepend = RS(11)
				'  response.write " codrel: " &  SCODREL & " master: " & smaster & " depen: " & sdepend & "<br>"

				sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL
				''   response.write sSQL
				Set rs2 = OFV.crearconsulta2(sSQL, cn)
				If Not rs2.EOF Then
					sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL & " AND (SITE = 0" & SESSION("SSITE") & " OR SITE = 0" & sSECSITE & ") "
					''     response.write sSQL
					Set rs2 = OFV.crearconsulta2(sSQL, cn)
					svasec = "NO"
					If Not rs2.EOF Then
						''   response.write " -codrel: " &  SCODREL & " site: " & rs2(2) & " site sel: " & SESSION("SSITE") &  "<br>"

						If CLng("0" & SESSION("SSITE")) = CLng("0" & sSECSITE) Then
							svasec = "OK"
						ElseIf CLng("0" & SESSION("SSITE")) = CLng("0" & rs2(2)) Then
							svasec = "OK"
						Else
							sSQL = "SELECT ER.* FROM ECOREL ER,ECORELSITE ES where ER.CODREL = ES.CODREL" & " AND ES.SITE = 0" & SESSION("SSITE") & " and ER.ECODEPEND = 0" & rs(0)
							''     response.write sSQL
							Set rs2 = OFV.crearconsulta2(sSQL, cn)
							If rs2.eof Then
								svasec = "OK"
							End If
						End If

						If svasec = "OK" Then
							xin = xin + 1
							xinw = xin
							shref = "OrganigramaUNgraf.asp?id=" & rs(0)
							wwwsettree = wwwsettree & "myTree.add2(" & xin & ",0,'" &Left( rs(1),120)& "',120,60,'" & session("CMFMENU") &"','#000000','" & shref & "',null,'1','NO');"
							sSQL = "SELECT DISTINCT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ER.CODREL,ER.ECOMASTER,ER.ECODEPEND "
							sSQL = sSQL & " FROM ECO E, ECOREL ER, ECO  ECO_1, segmento, ECOSITE ES "
							sSQL = sSQL & " where E.eco = ER.ECOMASTER and "
							sSQL = sSQL & " ER.ECODEPEND = ECO_1.eco and "
							sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and ES.ECO = ECO_1.ECO AND "
							sSQL = sSQL & " E.eco = " & RS(0)& " AND (ES.ECOSITE = 0" & SESSION("SSITE")
							sSQL = sSQL & " OR ES.ECOSITE = 0" & sSECSITE & ") and "
							sSQL = sSQL & " segmento.descri <> 'STAFF' order by ECO_1.ECONIV"

							'  response.write ssql
							Set RST1 = OFV.crearconsulta2(sSQL, cn)

							'Texto de Sectores Continuadores y Lineas
							Do While Not RST1.EOF


								xin = xin + 1
								xinw2=xin
								shref = "OrganigramaUNgraf.asp?id=" & RST1(0)
								wwwsettree = wwwsettree & "myTree.add2(" & xin & "," & xinw &",'" &Left(  RST1(1),120)& "',120,60,'" & session("CMFMENU") &"','#000000','" & shref & "',null,'1','NO');"

								sSQL = "SELECT DISTINCT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ER.CODREL,ER.ECOMASTER,ER.ECODEPEND "
								sSQL = sSQL & " FROM ECO E, ECOREL ER, ECO  ECO_1, segmento, ECOSITE ES "
								sSQL = sSQL & " where E.eco = ER.ECOMASTER and "
								sSQL = sSQL & " ER.ECODEPEND = ECO_1.eco and "
								sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and ES.ECO = ECO_1.ECO AND "
								sSQL = sSQL & " E.eco = " & RST1(0)& " AND (ES.ECOSITE = 0" & SESSION("SSITE")
								sSQL = sSQL & " OR ES.ECOSITE = 0" & sSECSITE & ") and "
								sSQL = sSQL & " segmento.descri <> 'STAFF' order by ECO_1.ECONIV"

								'  response.write ssql
								Set RST2 = OFV.crearconsulta2(sSQL, cn)

								'Texto de Sectores Continuadores y Lineas
								Do While Not RST2.EOF


									xin = xin + 1
                                   shref = "OrganigramaUNgraf.asp?id=" & RST2(0)
									wwwsettree = wwwsettree & "myTree.add2(" & xin & "," & xinw2 &",'" &Left(  RST2(1),120)& "',120,60,'" & session("CMFMENU") &"','#000000','" & shref & "',null,'1','NO');"


									RST2.MOVENEXT
								Loop
								RST1.MOVENEXT
							Loop


						End If
						'''    EXIT DO
					End If
				Else
					If SESSION("SSITE") = 1 Then
						xin = xin + 1
						ReDim Preserve vecsec(xin)
						ReDim Preserve vecsecd(xin)
						ReDim Preserve vecsecn(xin)
						vecsec(xin) = CLng(rs(0))
						vecsecd(xin) = rs(1)
						vecsecn(xin) = rs(5)
						'''    EXIT DO
					End If
				End If

				'''''''''''''






				rs.movenext
			Loop
			Call OFV.cerrarconsulta(rs)



			xinS = 0
			ReDim Preserve vecsecS(xinS)
			ReDim Preserve vecsecSD(xinS)
			ReDim Preserve vecsecSN(xinS)

			sSQL = "SELECT DISTINCT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ECOREL.CODREL FROM ECO, ECOREL, ECO  ECO_1, segmento, ECOSITE ES "
			sSQL = sSQL & "  where ECO.eco = ECOREL.ECOMASTER and "
			sSQL = sSQL & " ECOREL.ECODEPEND = ECO_1.eco and "
			sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and ES.ECO = ECO_1.ECO AND "
			sSQL = sSQL & " ECO.eco = " & nID & " AND (ES.ECOSITE = 0" & SESSION("SSITE")
			sSQL = sSQL & " OR ES.ECOSITE = 0" & sSECSITE & ") and "
			sSQL = sSQL & " segmento.descri = 'STAFF'"
			'RESPONSE.WRITE sSQL

			Set rs2 = OFV.crearconsulta2(sSQL, cn)
			Do Until rs2.EOF

				''''''''''''''''
				SCODREL = RS2(9)

				sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL
				'  response.write strsql
				Set rs3 = OFV.crearconsulta2(sSQL, cn)
				If Not rs3.EOF Then
					sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL & " AND (SITE = 0" & SESSION("SSITE") & " OR SITE = 0" & sSECSITE & ") "
					''     response.write sSQL
					Set rs3 = OFV.crearconsulta2(sSQL, cn)
					svasec = "NO"
					If Not rs3.EOF Then
						If CLng("0" & SESSION("SSITE")) = CLng("0" & sSECSITE) Then
							svasec = "OK"
						ElseIf CLng("0" & SESSION("SSITE")) = CLng("0" & rs3(2)) Then
							svasec = "OK"
						Else
							sSQL = "SELECT ER.* FROM ECOREL ER,ECORELSITE ES where ER.CODREL = ES.CODREL" & " AND ES.SITE = 0" & SESSION("SSITE") & " and ER.ECODEPEND = 0" & rs2(0)
							''     response.write sSQL
							Set rs3 = OFV.crearconsulta2(sSQL, cn)
							If rs3.eof Then
								svasec = "OK"
							End If
						End If

						If svasec = "OK" Then
							xinS = xinS + 1
							ReDim Preserve vecsecS(xinS)
							ReDim Preserve vecsecSD(xinS)
							ReDim Preserve vecsecSn(xinS)
							vecsecS(xinS) = CLng(rs2(0))
							vecsecSD(xinS) = rs2(1)
							vecsecSn(xinS) = rs2(5)
							'''  EXIT DO
						End If
					End If
				Else
					If SESSION("SSITE") = 1 Then
						xinS = xinS + 1
						ReDim Preserve vecsecS(xinS)
						ReDim Preserve vecsecSD(xinS)
						ReDim Preserve vecsecSn(xinS)
						vecsecS(xinS) = CLng(rs2(0))
						vecsecSD(xinS) = rs2(1)
						vecsecSn(xinS) = rs2(5)
						'''  EXIT DO
					End If
				End If

				'''''''''''''




				rs2.movenext
			Loop

			If UBound(vecsecS) > 0 Then
				'      tabp = tabp & GeneraStaff()
			End If

			If UBound(vecsec) > 0 Then
				'      tabp = tabp & GeneraResto()
			End If

			shtml = shtml & " <div id='myTreeContainer' style='position:relative;top:10px;left:1px;' > "
			shtml = shtml & " </div>"
			shtml = shtml & " <form id=ejecutar style='display:none' method=post></form>"
			shtml = shtml & "<script>"
			shtml = shtml & " var myTree = null;"
			delim = ""
			Amaxlevel = ""

			shtml=shtml &  "var aasis=new Array('-11');"


			shtml = shtml & " function CreateTree() { "
			shtml = shtml & "myTree = new ECOTree('myTree','myTreeContainer');"
			'''    shtml=shtml &  "myTree.config.colorStyle = ECOTree.CS_LEVEL;"
			shtml = shtml & "myTree.config.colorStyle = ECOTree.CS_NODE;"
			shtml = shtml & "myTree.config.nodeFill = ECOTree.NF_FLAT;"
			shtml = shtml & "myTree.config.useTarget = true;"
			shtml = shtml & "myTree.config.selectMode = ECOTree.SL_NONE;"
			shtml = shtml & "myTree.config.defaultNodeWidth = 85;"
			shtml = shtml & "myTree.config.defaultNodeHeight = 60;"
			shtml = shtml & "myTree.config.iSubtreeSeparation = 30;"
			shtml = shtml & "myTree.config.iSiblingSeparation = 20;"
			shtml = shtml & "myTree.config.iLevelSeparation = 50;"
			'shtml = shtml & "myTree.config.iRootOrientation = ECOTree.RO_LEFT;"
			' shtml = shtml & "myTree.config.nodeColor = " & Chr(34) & "#77AADD" & Chr(34) & ";"
			shtml = shtml & "myTree.config.topXAdjustment = 265;"
			shtml=shtml &  "myTree.config.rootXOffset =150; "
			shtml = shtml & "myTree.config.linkColor = 'black' ; "
			' shtml = shtml & "myTree.config.iNodeJustification =ECOTree.NJ_CENTER; "

			shtml = shtml & wwwsettree

			shtml = shtml & "myTree.UpdateTree();"
			shtml = shtml & "}"

			shtml = shtml & " var elem = document.getElementById('myTreeContainer');"
			'shtml = shtml & " var elemcc = document.getElementById('ECOTreecanvas');alert(elemcc.style.width);"
			shtml = shtml & "elemLeft = elem.offsetLeft; "
			shtml = shtml & "elemTop = elem.offsetTop;"
			shtml = shtml & " CreateTree();elements = []; "

			shtml = shtml & " for (var k = 0; k < myTree.nDatabaseNodes.length; k++)"
			shtml = shtml & "{  "
			shtml = shtml & "	node = myTree.nDatabaseNodes[k];"
			shtml = shtml & "	elements.push({width: node.w,"
			shtml = shtml & "	height: node.h,"
			shtml = shtml & "	top: node.YPosition,"
			shtml = shtml & "	 left: node.XPosition,"
			shtml = shtml & "	 url:node.target"
			shtml = shtml & "	});"
			shtml = shtml & "var g = document.createElement('div'); g.style.width= node.w + 'px'; g.style.height=node.h + 'px'; g.style.backgroundColor='';g.style.position='absolute';"
			shtml = shtml &  " g.style.left=elemLeft + node.XPosition + 'px'; g.style.top=elemTop + node.YPosition + 'px';"
			shtml = shtml & "g.id= k + 'cmd';"
			shtml= shtml  &  "g.addEventListener('mouseover', mouseOver (g), false);"
			shtml= shtml  &  "g.addEventListener('mouseout', mouseOut(g), false);"
			shtml= shtml  &  "g.onclick=function ejecf(){ var xxx= this.id.replace(/[^0-9]/g,'');var fx=document.getElementById('ejecutar'); fx.action=elements[xxx].url;fx.submit(); ;return false;};"
			shtml = shtml & "document.body.appendChild(g);"


			shtml = shtml & "}"

			shtml=shtml & " function mouseOver (e) {e.style.cursor='hand';e.style.cursor='pointer'};"
			shtml=shtml & " function mouseOut (e) {e.style.cursor='auto;'};"
			'shtml=shtml & " function ejecf (e) { alert(e);};"

			shtml = shtml & "    function mostrar(showHideDiv) { "
			shtml = shtml & "    	var ele = document.getElementById(showHideDiv);"

			shtml = shtml & "	    	if(ele.style.display == 'none') {"
			shtml = shtml & "	            ele.style.display = 'block';"

			shtml = shtml & "	      	}"
			shtml = shtml & "	    	else {"
			shtml = shtml & "	    		ele.style.display = 'none';"

			shtml = shtml & "	    } "
			shtml = shtml & "	    	}"


            shtml=shtml & "function largo()"
            shtml=shtml & "{"
         '    shtml=shtml & " var ele = document.getElementById('ECOTreecanvas');"
        '   shtml=shtml & "  ele.ssetAttribute('style','width: 500px;');"
           shtml=shtml & " }"

			shtml = shtml & "  </script>"




			XFECODSC = "N"

			sSQL = "SELECT * FROM ECODESCRIPCION WHERE ECO = 0" & nID
			Set rs2 = OFV.crearconsulta2(sSQL, cn)
			If Not rs2.EOF Then XFECODSC = "S"




			Tabla = Tabla & "<br><br><font face=verdana color=#000077 size=1>"


			sinput = sft2 & "SUCURSAL:" & sff
			celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
			sinput = sft & sSUCN & sSUCA & sff

			If XFECODSC = "N" Then
				celda = celda & OFV.gencelda("", "aaa COLSPAN=3 style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")
			Else
				celda = celda & OFV.gencelda("", "aaa  style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")

				sAccion = "genDSCSECTOR.asp?SECTOR=" & nID

				sHref   = sAccion
				sInput  = "<U>" & "DESCRIPCION" & "</U>"
				sInput  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",sInput,"si")

				celda = celda & OFV.gencelda("", "VT COLSPAN=2 bgcolor=#FACC66 title='VER DESCRIPCION DEL PUESTO.'  style='cursor:pointer;COLOR:#FFFFFF;border-style:solid;border-width:1;border-Color:#336699;'", "CENTER", "10%", "", "", "", sinput, "si")
			End If

			fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

			sinput = sft2 & "AREA:" & sff
			celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
			sinput = sft & sSector & sff
			celda = celda & OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")
			sinput = sft2 & "CODIGO:" & sff
			celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
			sinput = sft & sCodin & sff
			celda = celda & OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
			fila = FILA & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

			sinput = sft2 & "NIVEL:" & sff
			celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
			sinput = sft & sNIVELA & sff
			celda = celda & OFV.gencelda("", "aaa COLSPAN=3 style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")
			fila = FILA & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

			Tabla = Tabla & OFV.gentabla("", "", "", "100% bgcolor=#ffffff style='border-style:solid;border-width:1;border-Color:#cccccc;' ", "", "0", "", "1", "1", "", "", fila, "si")


			'   sSQL = "SELECT * FROM usuarios where sector = " & nID & " AND SITE2 = " & sSECSITE & " order by nivel "
			'   response.write ssql
			Dim nusuarios
			nusuarios=0
			'  Set rs = OFV.crearconsulta2(sSQL, cn2)
			'  response.write  nID & "--" &  sSECSITE
			Call getUsrInSect(nID,sSECSITE)
			arruid=Split(session("SGU_UIGuserid"),Chr(9))

			xineco = 0
			Strsql = "Select * FROM ECOPERFIL,tipoperfil WHERE ECOPERFIL.codperfil=tipoperfil.tipoperfil and  codsector = 0" & nID
			'response.write strsql
			Set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


			Do While Not rsx2.eof
				ReDim Preserve Vecoperfil(3,xineco)
				Vecoperfil(0,xineco)=rsx2(0)  ' objid
				Vecoperfil(1,xineco)=rsx2(2) ' codperfil
				Vecoperfil(2,xineco)=0
				Vecoperfil(3,xineco)=rsx2("descri")

				xineco=xineco+1
				rsx2.movenext
			Loop
			Call ofv.cerrarconsulta(rsx2)
			sImageva    = "&nbsp;&nbsp;" & "<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>"
			If session("SGU_UIGuserid") <> "" Then
				sinput = sft2 & "PERFIL/FUNCION" & sff
				celda =  OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")

				tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
				sinput = sft2 & "PERSONAL" & sff
				celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")
				'''   sinput = sft2 & "NIVEL/PERFIL" & sff


				fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
				Dim TXESTACC0,TXESTACC1,TXESTACC2
				TXESTACC0=0
				TXESTACC1=0
				TXESTACC2=0
				flag=False
				susuario=""






				For Iid=0 To UBound(arruid)
					'      Do Until rs.EOF
					celda=""
					'sinput = tipo & "&nbsp;" & sft & "<b>"  & rs(1) & "</b>" & sff
					usrx=getinfouser(arruid(Iid))
					sinput = tipo & "&nbsp;" & sft & "<b>"  & session("SGU_nombre") & "</b>" & sff
					If Not flag Then
						' capo=rs(1)
						capo=session("SGU_nombre")
						flag=True
					End If
					sinput2 = "<a   style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
					sinput2 = sinput2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
					'''sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "mostrarmenu('"& rs(0) &"','" & rs(1) & "');" & chr(34)
					sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' "
					sinput2 = sinput2  & "  ><b>" & sinput & "</b></a>"

					sinput2 = "<b>" & sinput & "</b>"

					celdaz = OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput2, "si")
					' XPUESTO = RS("NIVEL")
					XPUESTO=session("SGU_idnivel")
					sSQL = "SELECT * FROM NIVELCONFIDEN where NCCod = 0" & XPUESTO
					' response.write ssql
					Set rsX = OFV.crearconsulta2(sSQL, cn2)
					If Not rsX.EOF Then
						XPUESTO = RSX(1)
					Else
						XPUESTO = "---"
					End If

					''' esp blp ''''''
					XPUESTO = ""
					''' fin esp blp ''''''

					sImageva    = "&nbsp;&nbsp;" & "<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>"


					XPUESTOP = session("SGU_codigopuesto") 'RS("PUESTO")

					For Np=0 To xineco-1
						If CStr(Vecoperfil(1,np) &"")=CStr(XPUESTOP &"") Then
							Vecoperfil(2,np)=1

							Exit For
						End If
					Next


					'  RESPONSE.WRITE "CC=" & CC

					sSQL = "SELECT * FROM TIPOPERFIL where TIPOPERFIL = '" & XPUESTOP & "' "
					' response.write ssql
					Set rsX = OFV.crearconsulta2(sSQL, cn)
					If Not rsX.EOF Then


						XPUESTOP = "<B>" & rsX(1) & "</B>"


						Strsql = "Select * FROM ECOPERFIL WHERE "
						'Strsql = Strsql & " codsector = 0" & nID & " AND codperfil ='" & RS("PUESTO") & "' "
						Strsql = Strsql & " codsector = 0" & nID & " AND codperfil ='" & session("SGU_codigopuesto") & "' "
						'RESPONSE.WRITE STRSQL & "<BR>"
						Set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not rsX2.EOF Then

							sAccion = "genPTOFUN.asp?PUESTO=" & rsx2(0)

							sstilo2 = "  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
							sstilo2 = sstilo2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
							sstilo2 = sstilo2  & "none" & Chr(34) & ";' "
							'
							xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
							sHref   = sAccion
							sInput  = sft & "&nbsp;" & XPUESTOP & sff & sImageva
							Celdaobs  = ofv.GenCelda("","ut" & xsffun & sstilo2 & " style='cursor:pointer;border-style:solid;border-width:1;border-Color:#336699;' title='VER DESCRIPCION DEL PUESTO.' ","","","1","","",sInput,"si")
							Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",Celdaobs,"si")

							celda = celda    & Celdaobs & celdaz

						Else

							sinput =  sft & "&nbsp;" & XPUESTOP & sff
							celda = celda &  OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput, "si") & celdaz

						End If

					Else
						XPUESTOP = "---"

						sinput =  sft & "&nbsp;" & XPUESTOP & sff
						celda = celda &  OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput, "si") & celdaz

					End If





					fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
					'        rs.movenext
					tipo="<b><font face=corflaticon size=2 color=#6699cc >=</font></b>"

					'     Loop
				Next

				For Np=0 To xineco-1
					If  Vecoperfil(2,np)=0  Then
						sAccion = "genPTOFUN.asp?PUESTO=" & Vecoperfil(0,NP)

						sstilo2 = "  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
						sstilo2 = sstilo2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
						sstilo2 = sstilo2  & "none" & Chr(34) & ";' "
						'
						xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
						sHref   = sAccion
						sInput  = sft & "&nbsp;<b>" & Vecoperfil(3,NP) &"</b>" & sff & sImageva
						Celdaobs  = ofv.GenCelda("","ut" & xsffun & sstilo2 & " style='cursor:pointer;border-style:solid;border-width:1;border-Color:#336699;' title='VER DESCRIPCION DEL PUESTO.' ","","","1","","",sInput,"si")
						Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",Celdaobs,"si")
						celdaz = OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", "---", "si")
						celda =  Celdaobs & celdaz

						fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
						'        rs.movenext
						tipo="<b><font face=corflaticon size=2 color=#6699cc >=</font></b>"
					End If
				Next
				'total
				tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
				sinput =  "&nbsp;"
				celda =  OFV.gencelda("", "aaa colspan=2 bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "50%", "", "", "", sinput, "si")
				'  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
				Tabla = Tabla & "<br>" & OFV.gentabla("", "", "", "100%  bgcolor=#ffffff style='border-style:solid;border-width:1;border-Color:#cccccc;' ", "", "0", "", "1", "1", "", "", fila, "si")




			Else

				If xineco > 0 Then
					sinput = sft2 & "PERFIL/FUNCION" & sff
					celda =  OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")

					tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
					sinput = sft2 & "PERSONAL" & sff
					celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")
					fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

					For Np=0 To xineco-1
						If  Vecoperfil(2,np)=0  Then
							sAccion = "genPTOFUN.asp?PUESTO=" & Vecoperfil(0,NP)

							sstilo2 = "  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
							sstilo2 = sstilo2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
							sstilo2 = sstilo2  & "none" & Chr(34) & ";' "
							'
							xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
							sHref   = sAccion
							sInput  = sft & "&nbsp;<b>" & Vecoperfil(3,NP) &"</b>" & sff & sImageva
							Celdaobs  = ofv.GenCelda("","ut" & xsffun & sstilo2 & " style='cursor:pointer;border-style:solid;border-width:1;border-Color:#336699;' title='VER DESCRIPCION DEL PUESTO.' ","","","1","","",sInput,"si")
							Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",Celdaobs,"si")
							celdaz = OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", "---", "si")
							celda =  Celdaobs & celdaz

							fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
							'        rs.movenext
							tipo="<b><font face=corflaticon size=2 color=#6699cc >=</font></b>"
						End If
					Next
					'total
					tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
					sinput =  "&nbsp;"
					celda =  OFV.gencelda("", "aaa colspan=2 bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "50%", "", "", "", sinput, "si")
					'  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
					Tabla = Tabla & "<br>" & OFV.gentabla("", "", "", "100%  bgcolor=#ffffff style='border-style:solid;border-width:1;border-Color:#cccccc;' ", "", "0", "", "1", "1", "", "", fila, "si")




				End If

			End If





			tabp = tabp & OFV.GenFcelda()
			tabp = tabp & OFV.GenFRow()
			tabp = tabp & OFV.genFtabla()



			sHTML = sHTML &  tabp & "</font>"


		Else


		End If

		Call OFV.cerrarconn(cn)
		Call OFV.cerrarconn(cn2)
		Call OFV.cerrarconn(cn3)
		If request.querystring("mseg")="M" Then
			sHTML = sHTML  & "</body></html>"
		Else

			sHTML = sHTML  & "<br></body></html>"
		End If


		Generar = sHTML

	End If





End If

Response.Write sHTML


%>