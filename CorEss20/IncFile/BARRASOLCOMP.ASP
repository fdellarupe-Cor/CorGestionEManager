Function armasolapaComp(sCodigo,sTipo,sFuncionX)
	Dim sHTML
	Dim Celda
	bModificar = False
	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False
    If sFuncion <> "" Then sFuncion=sFuncionX
	sTipo     = request.querystring("Tipo")
	sCodigo   = request.querystring("Codigo")
	'sFuncion  = request.querystring("Funcion")
	sSClick   = request.querystring("SSClick")
	sMenu     = request.querystring("Menu")
	sSector   = request.querystring("Sector")
	sNIV1     = request.querystring("NIV1")
	sNIV2     = request.querystring("NIV2")
	sNIV3     = request.querystring("NIV3")
	sNIV4     = request.querystring("NIV4")
	sNIV5     = request.querystring("NIV5")

	sPerfil      = request.querystring("Perfil")
	sComplemento = request.querystring("Complemento")
	sComplemento = ofv.convertircar(sComplemento,"_"," ")

	''  response.write "recibe " & request.querystring & "<br>"

	wsSClick = "V" & sSClick

	If  sSClick = "" Then     sSClick      = 99


	Set cn = ofv.conectar(ofv.strconn2)



	'sHTML = "<html>"

	wxs = "<script language=JavaScript>" & Chr(13)
	wxs = wxs & " var Fbtini = 0;" & Chr(13)
	wxs = wxs & " function IniWin() { " & Chr(13)
	wxs = wxs & "  if (document.all.SSOLACT) { " & chr(13)
	wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
	wxs = wxs & " } else { " & chr(13)
	wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
	wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
	wxs = wxs & "  } " & chr(13)
	wxs = wxs & " } " & Chr(13)
	wxs = wxs & " </script>" & Chr(13)

	'sHTML = sHTML & wxs
   'ver onload
	'sHTML = sHTML & "<body xonload='IniWin();' >"

    sHTML=""
	Dim stabla(20)
	Dim solord(20)

	solord(0) = 0
	solord(1) = 15
	solord(2) = 1
	solord(3) = 6
	solord(4) = 3
	solord(5) = 2
	solord(6) = 4
	solord(7) = 5
	solord(8) = 7
	solord(9) = 8
	solord(10) = 11
	solord(11) = 10
	solord(12) = 9
	solord(13) = 12
	solord(14) = 13
	solord(15) = 14
	solord(16) = 16
	solord(17) = 17
	solord(18) = 18
	solord(19) = 19
	solord(20) = 20






	Dim Nombres(20)

	Dim ASPs(20)

	ASPs(0) = "imagencont2.asp?skind=1"


	If sFuncion = "E" Then
		ASPs(1) = "edGenKomp.asp?skind=1"
	Else
		ASPs(1) = "imagencontG.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(2) = "edEstrKomp.asp?skind=1"
	Else
		ASPs(2) = "imagencont3.asp?skind=1"
	End If

	ASPs(3) = "N"

	If sFuncion = "E" Then
		ASPs(4) = "edImagen2.asp?skind=1"
	Else
		ASPs(4) = "imagencont2.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(5) = "edCompuesto2.asp?skind=1"
	Else
		ASPs(5) = "ImagenIndex.asp?skind=1"
	End If




	If sFuncion = "E" Then
		ASPs(6) = "edimgarc2.asp?skind=1"
	Else
		ASPs(6) = "imagencont4.asp?skind=1"
	End If



	ASPs(7)  = "N"
	ASPs(8)  = "N"
	ASPs(9)  = "N"
	ASPs(10) = "N"
	ASPs(11) = "N"
	ASPs(12) = "N"
	ASPs(13) = "N"
	ASPs(14) = "N"
	ASPs(15) = "N"

	perfilsp = "S"







	siconmas = "ICorAcciones.jpg"



	Dim Solapas(20)
	Dim Solapas2(20)

	For z=0 To 20
		Solapas(z) = "N"
	Next



	Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


	If Session("Segco") <> "S" Then
		For xx = 1 To 20
			Nombres(xx) = "N"
		Next

		Solapas(15)  = REPLACE(RS(31)," ","*")
		Solapas2(15) = ASPs(15)
		xstrsol      = 1
		xfinal       = 1
		sSClick      = 99


	Else

		For xx = 1 To 20
			solord(xx) = 17
			Nombres(xx) = "N"
		Next



		For xx = 2 To 16
			stabla (xx - 1) = rs(xx + 30)
		Next



		yy = 0
		For xx = 1 To 15

			If cint(stabla (xx)) < 99 Then
				yy = yy + 1
				solord (stabla (xx)) = xx
			End If
		Next
		tyy = yy




		If session("ModMain") = "WK"  Then
			Solapas(0)  = "WORKFLOW"
			Solapas2(0) = ASPs(0)
			xstrsol     = 0
		Else
			xstrsol     = 1
		End If




		X=0
		For z = 2 To 16

			If rs(z) Then

				If (sFuncion = "E" And z <> 4  ) Or (sFuncion = "C" And z <> 21 And z <> 4 And z <> 61 )  Then
					x = x + 1
					Solapas(z-1)  = REPLACE(RS(z + 15)," ","*")
					Solapas2(z-1) = ASPs(z-1)




				End If
			End If

		Next


		xfinal = ubound(Solapas)

	End If

	''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

	Call ofv.cerrarconsulta(rs)
	Dim FTABELENUL
	'FTABELENUL= "<div id=menucor data-cli='"&session("APLMDCLI")& "' ><nav><h2><i class='la la-bars'></i>SOLAPAS</h2><UL>"

	FTABELENUL="<div id=contm STYLE='overflow:hidden;' xclass=sidebar><ul class=ulmenu STYLE= 'overflow:auto;Height:100%;' id=myMenu >" & chr(13)
	FTABELENUL=FTABELENUL & chr(13) & "<li><div class='item-container'><div class='item-text'><img src='" & Session("CliLogo") &"'><hr></div></div></li>" & Chr(13)

	xSolasp = ""
	xSolasp = xSolasp & "BarraCompSol.asp?Tipo=" & sTipo & "&Codigo="
	xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector
	xSolasp = xSolasp & "&Funcion=" & sFuncion
	xSolasp = xSolasp & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
	xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
	xSolasp = xSolasp & "&SSCLICK="

	r1 = ""


	If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
		xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
		xGenConta = ofv.FTitDat1

		xsclase = "xts"

	Else



		xsclase = "xts"

	End If

	xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
	xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"

	wsSClick2 = sSClick

	xsolini = 99


	For x = xstrsol To xfinal
		If solapas(solord(x)) <> "N" Then
			sSolapa =  solapas(solord(x))

			If  xsolini = 99 Then  xsolini = solord(x)

			If  sSClick = 99 Then  sSClick = xsolini

			If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


				sInput  = replace(Ucase(sSolapa),"*"," ") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"


			Else

				sInput  = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"

			End If


			If session("WEBDESK") = "S" Then


				sInput  = "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & "</font>"
				sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")
				' FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(sSolapa)  & "</a></font></LI>"

			'	FTABELENUL=FTABELENUL & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 HREF='"& xSolasp & solord(x) & "'> "  &  UCase(sSolapa) & "</a></font></LI>"
				FTABELENUL=FTABELENUL & "<li><div class='item-container'><div class='item-text' data-sol='" & Capitalize(sSolapa) & "'><a class=ntmnn target=Principal data-id='s1solapa" & solord(x) & "' class=""link"" href='"& xSolasp & solord(x) & "'>"  &  Capitalize(sSolapa)  & "</a></div></div></li>"

				Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

				If (solord(x)) = 0 Then
					wwwestilo = " style='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
				End If


				CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
				FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



				If sFuncion = "E" Or sFuncion = "ED"  Then

				Else

				End If


			Else


				t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
				t3 = t3 & sInput
				t3 = t3 & "</A>"



				t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

				r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

			End If

			If clng(sSClick) = clng(solord(x)) Then

				xverasp = solapas2(solord(x)) & "&Tipocomp=" & sTipo & "&Codcomp=" & sCodigo & "&id=" & sCodigo
				xverasp = xverasp & "&Menu=" & sMenu & "&Sector=" & sSector
				xverasp = xverasp & "&GLBSOLN=S" & solord(x)
				xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

				sHTML1 = "" & chr(13)
				sHTML1 = sHTML1 & "<form name=SSOLACT action='" & xverasp & "' "
				sHTML1 = sHTML1 & " method=post target='Principal' style='display:none;' >" & chr(13)
				sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
				sHTML1 = sHTML1 & "</form>" & chr(13)

				sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
				sHTML1 = sHTML1 & " window.top.CorSol='" & replace(Ucase(sSolapa),"*"," ") & "'; " & Chr(13)
				sHTML1 = sHTML1 & "</script>" & Chr(13)


				If session("WEBDESK") = "S" Then




				Else

					If  Session("FULLGUI") = "N" Then
						styl = "background-color:#888888;"
					Else
						styl = XspEx("#888888","#cccccc",0) & ";"
					End If

					xlfun = " onclick=""document.all.MDBEDL.style.display='block'; document.all.DBEDTD1.style.display='none'; document.all.DBEDTD2.style.display='block'; document.all.MM2.style.display='block'; "" "
					xlfun2 = " onclick=""document.all.MDBEDL.style.display='none'; document.all.DBEDTD1.style.display='block'; document.all.DBEDTD2.style.display='none'; document.all.MM2.style.display='none'; "" "

					macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

					xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


					macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
					macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

					marow = ofv.GenRow("","","","","20","","",macel,"si")

					sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

				End If

			End If


		End If
	Next

	If session("WEBDESK") = "S" Then


		'''    sHTML = sHTML & Celda
		'''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

		'sHTML = sHTML & "<div id='WKTBSP1'  >"
	'	sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "aaa  ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN, "si")
		'sHTML = sHTML & "</div>"

	Else


		'sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"

		If  Session("FULLGUI") = "N" Then
			styl = "background-color:#888888;"
		Else
			styl = XspEx("#888888","#cccccc",0) & ";"
		End If

		styl = ""

		If ofv.ModTitCls = "1" Then

		'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

			''    sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,-2,-2,"","",r1,"si")

		Else

		'	sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			''   sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,2,2,"","",r1,"si")

		End If

	'	sHTML = sHTML & "</div>"

	End If
	FTABELENUL=FTABELENUL &"</UL></DIV>"
	armasolapaComp= FTABELENUL

	Call ofv.cerrarconn(cn)

	'sHTML = sHTML & "</body></html>"





End Function


Function armasolapaComp2ww(sCodigo,sTipo,sFuncionX)
	Dim sHTML
	Dim Celda
	bModificar = False
	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False

	sTipo     = request.querystring("Tipo")
	sCodigo   = request.querystring("Codigo")
	sFuncion  = request.querystring("Funcion")
	sSClick   = request.querystring("SSClick")
	sMenu     = request.querystring("Menu")
	sSector   = request.querystring("Sector")
	sNIV1     = request.querystring("NIV1")
	sNIV2     = request.querystring("NIV2")
	sNIV3     = request.querystring("NIV3")
	sNIV4     = request.querystring("NIV4")
	sNIV5     = request.querystring("NIV5")

	sPerfil      = request.querystring("Perfil")
	sComplemento = request.querystring("Complemento")
	sComplemento = ofv.convertircar(sComplemento,"_"," ")

	''  response.write "recibe " & request.querystring & "<br>"

	wsSClick = "V" & sSClick

	'If  sSClick = "" Then     sSClick      = 99


	Set cn = ofv.conectar(ofv.strconn2)



	sHTML = "<html>"

	wxs = "<script language=JavaScript>" & Chr(13)
	wxs = wxs & " var Fbtini = 0;" & Chr(13)
	wxs = wxs & " function IniWin() { " & Chr(13)
	wxs = wxs & "  if (document.all.SSOLACT) { " & chr(13)
	wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
	wxs = wxs & " } else { " & chr(13)
	wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
	wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
	wxs = wxs & "  } " & chr(13)
	wxs = wxs & " } " & Chr(13)
	wxs = wxs & " </script>" & Chr(13)

	sHTML = sHTML & wxs

	sHTML = sHTML & "<body onload='IniWin();' >"


	Dim stabla(20)
	Dim solord(20)

	solord(0) = 0
	solord(1) = 15
	solord(2) = 1
	solord(3) = 6
	solord(4) = 3
	solord(5) = 2
	solord(6) = 4
	solord(7) = 5
	solord(8) = 7
	solord(9) = 8
	solord(10) = 11
	solord(11) = 10
	solord(12) = 9
	solord(13) = 12
	solord(14) = 13
	solord(15) = 14
	solord(16) = 16
	solord(17) = 17
	solord(18) = 18
	solord(19) = 19
	solord(20) = 20






	Dim Nombres(20)

	Dim ASPs(20)

	ASPs(0) = "imagencont2.asp?skind=1"


	If sFuncion = "E" Then
		ASPs(1) = "edGenKomp.asp?skind=1"
	Else
		ASPs(1) = "imagencontG.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(2) = "edEstrKomp.asp?skind=1"
	Else
		ASPs(2) = "imagencont3.asp?skind=1"
	End If

	ASPs(3) = "N"

	If sFuncion = "E" Then
		ASPs(4) = "edImagen2.asp?skind=1"
	Else
		ASPs(4) = "imagencont2.asp?skind=1"
	End If

	If sFuncion = "E" Then
		ASPs(5) = "edCompuesto2.asp?skind=1"
	Else
		ASPs(5) = "ImagenIndex.asp?skind=1"
	End If




	If sFuncion = "E" Then
		ASPs(6) = "edimgarc2.asp?skind=1"
	Else
		ASPs(6) = "imagencont4.asp?skind=1"
	End If



	ASPs(7)  = "N"
	ASPs(8)  = "N"
	ASPs(9)  = "N"
	ASPs(10) = "N"
	ASPs(11) = "N"
	ASPs(12) = "N"
	ASPs(13) = "N"
	ASPs(14) = "N"
	ASPs(15) = "N"

	perfilsp = "S"







	siconmas = "ICorAcciones.jpg"



	Dim Solapas(20)
	Dim Solapas2(20)

	For z=0 To 20
		Solapas(z) = "N"
	Next



	Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


	If Session("Segco") <> "S" Then
		For xx = 1 To 20
			Nombres(xx) = "N"
		Next

		Solapas(15)  = REPLACE(RS(31)," ","*")
		Solapas2(15) = ASPs(15)
		xstrsol      = 1
		xfinal       = 1
		sSClick      = 99


	Else

		For xx = 1 To 20
			solord(xx) = 17
			Nombres(xx) = "N"
		Next



		For xx = 2 To 16
			stabla (xx - 1) = rs(xx + 30)
		Next



		yy = 0
		For xx = 1 To 15

			If cint(stabla (xx)) < 99 Then
				yy = yy + 1
				solord (stabla (xx)) = xx
			End If
		Next
		tyy = yy




		If session("ModMain") = "WK"  Then
			Solapas(0)  = "WORKFLOW"
			Solapas2(0) = ASPs(0)
			xstrsol     = 0
		Else
			xstrsol     = 1
		End If




		X=0
		For z = 2 To 16

			If rs(z) Then

				If (sFuncion = "E" And z <> 4  ) Or (sFuncion = "C" And z <> 21 And z <> 4 And z <> 61 )  Then
					x = x + 1
					Solapas(z-1)  = REPLACE(RS(z + 15)," ","*")
					Solapas2(z-1) = ASPs(z-1)




				End If
			End If

		Next


		xfinal = ubound(Solapas)

	End If

	''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

	Call ofv.cerrarconsulta(rs)
	Dim FTABELENUL
	FTABELENUL= "<div style='display:none' id=menucor data-cli='"&session("APLMDCLI")& "' ><nav><h2><i class='la la-bars'></i>SOLAPAS</h2><UL>"
	xSolasp = ""
	xSolasp = xSolasp & "BarraCompSol.asp?Tipo=" & sTipo & "&Codigo="
	xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector
	xSolasp = xSolasp & "&Funcion=" & sFuncion
	xSolasp = xSolasp & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
	xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
	xSolasp = xSolasp & "&SSCLICK="

	r1 = ""


	If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
		xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
		xGenConta = ofv.FTitDat1

		xsclase = "xts"

	Else



		xsclase = "xts"

	End If

	xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
	xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"

	wsSClick2 = sSClick

	xsolini = 99


	For x = xstrsol To xfinal
		If solapas(solord(x)) <> "N" Then
			sSolapa =  solapas(solord(x))

			If  xsolini = 99 Then  xsolini = solord(x)

			If  sSClick = 99 Then  sSClick = xsolini

			If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


				sInput  = replace(Ucase(sSolapa),"*"," ") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"


			Else

				sInput  = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
				xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"

			End If


			If session("WEBDESK") = "S" Then


				sInput  = "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & "</font>"
				sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")
				' FTABELENUL=FTABELENUL & "<LI data-id='s1solapa" & solord(x) & "' class=cn1 ><font color=  #fdfefe   size=1 face=verdana ><a  data-id='s1solapa" & solord(x) & "' class=""link"" xhref='#s1solapa" & solord(x) & "'> "  &  UCase(sSolapa)  & "</a></font></LI>"
				FTABELENUL=FTABELENUL & "<LI><font color=  #fdfefe   size=1 face=verdana ><a target=WK1 HREF='"& xSolasp & solord(x) & "'> "  &  UCase(sSolapa) & "</a></font></LI>"
				Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

				If (solord(x)) = 0 Then
					wwwestilo = " style='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
				End If


				CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
				FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



				If sFuncion = "E" Or sFuncion = "ED"  Then

				Else

				End If


			Else


				t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
				t3 = t3 & sInput
				t3 = t3 & "</A>"



				t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

				r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

			End If

			If clng(sSClick) = clng(solord(x)) Then

				xverasp = solapas2(solord(x)) & "&Tipocomp=" & sTipo & "&Codcomp=" & sCodigo & "&id=" & sCodigo
				xverasp = xverasp & "&Menu=" & sMenu & "&Sector=" & sSector
				xverasp = xverasp & "&GLBSOLN=S" & solord(x)
				xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

				sHTML1 = "" & chr(13)
				sHTML1 = sHTML1 & "<form name=SSOLACT action='" & xverasp & "' "
				sHTML1 = sHTML1 & " method=post target='Principal' style='display:none;' >" & chr(13)
				sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
				sHTML1 = sHTML1 & "</form>" & chr(13)

				sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
				sHTML1 = sHTML1 & " window.top.CorSol='" & replace(Ucase(sSolapa),"*"," ") & "'; " & Chr(13)
				sHTML1 = sHTML1 & "</script>" & Chr(13)


				If session("WEBDESK") = "S" Then




				Else

					If  Session("FULLGUI") = "N" Then
						styl = "background-color:#888888;"
					Else
						styl = XspEx("#888888","#cccccc",0) & ";"
					End If

					xlfun = " onclick=""document.all.MDBEDL.style.display='block'; document.all.DBEDTD1.style.display='none'; document.all.DBEDTD2.style.display='block'; document.all.MM2.style.display='block'; "" "
					xlfun2 = " onclick=""document.all.MDBEDL.style.display='none'; document.all.DBEDTD1.style.display='block'; document.all.DBEDTD2.style.display='none'; document.all.MM2.style.display='none'; "" "

					macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

					xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


					macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
					macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

					marow = ofv.GenRow("","","","","20","","",macel,"si")

					sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

				End If

			End If


		End If
	Next

	If session("WEBDESK") = "S" Then


		'''    sHTML = sHTML & Celda
		'''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

		sHTML = sHTML & "<div id='WKTBSP1'  >"
		sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "aaa  ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN, "si")
		sHTML = sHTML & "</div>"

	Else


		sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"

		If  Session("FULLGUI") = "N" Then
			styl = "background-color:#888888;"
		Else
			styl = XspEx("#888888","#cccccc",0) & ";"
		End If

		styl = ""

		If ofv.ModTitCls = "1" Then

			sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

			''    sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,-2,-2,"","",r1,"si")

		Else

			sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
			''   sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,2,2,"","",r1,"si")

		End If

		sHTML = sHTML & "</div>"

	End If
	FTABELENUL=FTABELENUL &"</UL></NAV></DIV>"
	armasolapaComp2= FTABELENUL

	Call ofv.cerrarconn(cn)

	'sHTML = sHTML & "</body></html>"





End Function