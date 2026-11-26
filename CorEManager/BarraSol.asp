<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	bModificar = False
	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	session("SETIDEMPRE")=""

	If session("APLMDCLI")="" Then
		aplcheck()
	End If

	If session("APLMDCLI")="BGAL" Then
		lengtd=""
	Else
		lengtd="20"
	End If
	iitemC=0

	If session("Nomodificar") = "S" Then bModificar = False

	sTipo     = request.querystring("Tipo")
	sCodigo   = request.querystring("Codigo")
	sFuncion  = request.querystring("Funcion")
	sFuncionSol  = request.querystring("Funcion")
	sSClick   = request.querystring("SSClick")
	sMenu     = request.querystring("Menu")
	sSector   = request.querystring("Sector")
	sNIV1     = request.querystring("NIV1")
	sNIV2     = request.querystring("NIV2")
	sNIV3     = request.querystring("NIV3")
	sNIV4     = request.querystring("NIV4")
	sNIV5     = request.querystring("NIV5")
	sMAINAPL  = request.querystring("MAINAPL")

	''Nuevo para Test
	sTest     = request.querystring("Test")

	' response.write sTest & " &Test "

	''Nuevo para Test

'	response.write "recibe " & request.querystring & "  "  & time() & "<br>"

	Set cn = ofv.conectar(ofv.strconn2)
	Strsql = "Select TIPOCPT From TIPOBJETO Where TIPOBJETO ='" & sTipo & "'"
	Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	If Not rsx1.eof Then
		TIPOCPT=rsx1(0)
	Else
		TIPOCPT="D"
	End If
	Call ofv.cerrarconsulta(rsx1)


	sHTML = "<html>"
	sHTML = sHTML & "<HEAD><META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " content='text/html; charset=" & XLCHARSET & "'></HEAD>"

	wxs = "<script language=JavaScript>" & Chr(13)
	wxs = wxs & " var Fbtini = 0;" & Chr(13)
	wxs = wxs & " function IniWin() { " & Chr(13)
	wxs = wxs & "  if (document.getElementById('SSOLACT')) { " & Chr(13)
	wxs = wxs & " document.getElementById('SSOLACT').submit(); " & Chr(13)
	wxs = wxs & " } else { " & Chr(13)
	wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & Chr(13)
	wxs = wxs & " setTimeout (""IniWin()"", 100); } " & Chr(13)
	wxs = wxs & "  } " & Chr(13)
	wxs = wxs & " } " & Chr(13)
	wxs = wxs & " </script>" & Chr(13)

	sHTML = sHTML & wxs
	shtml =shtml &  CorStyler2()
	  Whtml = " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/MENU.asp"">"
	    Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/line-awesome.css"">"
		Whtml = Whtml & " <script src='../v48/line-awesome/jquery.min.js'></script>"
	    Whtml = Whtml & " <script src='../v48/line-awesome/jquery.balloon.js'></script>"
	    Whtml = Whtml & " <script src='../v48/line-awesome/jqex.js'></script>"
    Whtml = Whtml & " <script src='../v48/line-awesome/jqestilo.js'></script>"
    sHTML= shtml & Whtml
	sHTML = sHTML & "<body class=Stbody id=bodyM name=bodyM onload='IniWin();' ><div id=conteM>"



	If sFuncion = "E" Then
		smodosol = "SP"
	Else
		smodosol = "TB"
	End If


	Dim stabla(20)
	Dim solord(20)
	Dim Nombres(20)
	Dim ASPs(20)
	Dim Solapas(20)
	Dim Solapas2(20)


	ASPs(0) = "CoWKFselect.asp?SKIND=1"

	wwwsaltador = "<BR>"


	If SMAINAPL = "OR" Or SMAINAPL = "PL" Then



		ASPs(0) = "COWKFVERIFICACION.asp?SKIND=1"

		solord(0) = 0
		solord(1) = 1
		solord(2) = 2
		solord(3) = 3
		solord(4) = 4
		solord(5) = 5
		solord(6) = 6
		solord(7) = 7
		solord(8) = 8
		solord(9) = 9
		solord(10) = 10
		solord(11) = 11
		solord(12) = 12
		solord(13) = 13
		solord(14) = 14
		solord(15) = 15
		solord(16) = 16
		solord(17) = 17
		solord(18) = 18
		solord(19) = 19
		solord(20) = 20



		perfilsp = "S"

		If SMAINAPL = "OR" Then
			sEsquema = "CIRCUITO AUTOEVALUACION"
		Else
			sEsquema = "PLAN DE CONTINGENCIAS"
		End If

		xestado = 1

		Strsql = "Select estado,CODOBJETO,TIPOBJETO,RELLINK From ORRSKCIRCUITO Where OBJID = 0" & sCodigo
		Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsx1.eof Then
			xestado = rsx1(0)
			sCodigoVINC = rsx1(1)
			sTIPOVINC = rsx1(2)
			sRELVINC = rsx1(3)
		End If
		Call ofv.cerrarconsulta(rsx1)


		If (sFuncion = "CD") Then
			ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=0"
		Else
			ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=1"
		End If




		For z=0 To 20
			Solapas(z) = "N"
		Next


		For xx = 1 To 20
			Nombres(xx) = "N"
		Next

		'Solapas(1)  = REPLACE("DETALLE DEL CIRCUITO"," ","*")
		Solapas(1)  = Replace("MATRIZ DE AUTOEVALUACION"," ","*")
		Solapas2(1) = ASPs(1)
		xstrsol      = 1
		xfinal       = 1
		'''   sSClick      = 99



		If session("ModMain") = "WK"  Then
			Solapas(0)  = "WORKFLOW"
			Solapas2(0) = ASPs(0)
			xstrsol     = 0

		Else
			xstrsol     = 1
		End If



		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Else


		ASPs(1) = "edPropiedades.asp?SKIND=1"

		If sFuncion = "E" Then
			''' ASPs(2) = "edConsideraciones.asp?SKIND=1"
			ASPs(2) = "EDNWEDT.asp?SKIND=1"
		Else
			'''    ASPs(2) = "CoConsideraciones.asp?SKIND=1"
			ASPs(2) = "CONWEDT.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(3) = "edSectores.asp?SKIND=1"
		Else
			ASPs(3) = "CoSectores.asp?SKIND=1"
		End If



		ASPs(4) = "edEventos.asp?SKIND=1"

		If sFuncion = "E" Then
			ASPs(5) = "edProcedimientos.asp?SKIND=1"
		Else
			ASPs(5) = "CoProcedimientos.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
		Else
			ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(6) = "edObjDsc.asp?SKIND=1"
		Else
			ASPs(6) = "CoObjDsc.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(7) = "edControles.asp?SKIND=1"
		Else
			ASPs(7) = "CoControles.asp?SKIND=1"
		End If
		If sFuncion = "E" Then
			ASPs(8) = "edEsquemas.asp?SKIND=1"
		Else
			ASPs(8) = "CoEsquemas.asp?SKIND=1"
		End If
		If sFuncion = "E" Then
			ASPs(9) = "edTablas.asp?SKIND=1"
		Else
			ASPs(9) = "CoTablasCons.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(10) = "edarchivos.asp?SKIND=1"
		Else
			ASPs(10) = "CoArchCons.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(11) = "edestr.asp?SKIND=1"
		Else
			''  ASPs(11) = "Coestrex.asp?SKIND=1"
			ASPs(11) = "CoestrCons.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(12) = "CursogramaLnch.asp?SKIND=1"
			'''   ASPs(12) = "edProcedimientos.asp?SKIND=1"
		Else
			ASPs(12) = "CursogramaLnch.asp?SKIND=1"
			''''' ASPs(12)="CursogramaView.asp?SKIND=2"
		End If

		sMAINAPL = request.querystring("MAINAPL")
		If sFuncion = "E" Then
			ASPs(13) = "edarchivos.asp?SKIND=2"
		Else
			ASPs(13) = "CoArchCons.asp?SKIND=2"
		End If

		If sFuncion = "E" Then
			ASPs(14) = "edFAQ.asp?SKIND=1"
		Else
			ASPs(14) = "CoFAQ.asp?SKIND=1"
		End If


		If sFuncion = "E" Then
			ASPs(15) = "edGeneral.asp?SKIND=1"
		Else


			If TIPOCPT="K" Then
				ASPs(15) = "CoGeneralForm.asp?SKIND=1"
			Else
				ASPs(15) = "CoGeneral.asp?SKIND=1"
			End If
			' ASPs(15) = "CoGeneral.asp?SKIND=1"
		End If

		If sFuncion = "E" Then
			ASPs(20) = "edportada.asp?SKIND=1"
		Else
			ASPs(20) = "Coportada.asp?SKIND=1"
		End If



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




		perfilsp = "S"


		siconmas = "ICorAcciones.jpg"
		Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs.eof Then
			sEsquema = rs(1)
			sPerfil = rs(3)
			COMENTARIOS=RS("COMENTARIOS")
		End If

		Call ofv.cerrarconsulta(rs)


		For z=0 To 20
			Solapas(z) = "N"
		Next




		Strsql = "Select * From Objperfildoc Where codobjeto =" & sCodigo & " and tipobjeto ='" & stipo & "' "
		'   response.write strsql
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If rs.eof Then
			perfilsp = "N"
		End If

		If perfilsp = "N" Then
			Call ofv.cerrarconsulta(rs)
			Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"

			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		End If
		'response.write Strsql
		If Session("Segco") <> "S" Then
			For xx = 1 To 20
				Nombres(xx) = "N"
			Next
         '   IF RS.EOF THEN
			Solapas(15)  ="S/D"' Replace(RS(31)," ","*")
		'	ELSE
		'	Solapas(15)  = Replace(RS(31)," ","*")
		'	END IF
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

				If CInt(stabla (xx)) < 99 Then
					yy = yy + 1
					solord (stabla (xx)) = xx
				End If
			Next
			tyy = yy



			'''  if session("ModMain") = "WK" and sFuncion = "E" then
			If session("ModMain") = "WK"  Then
				'  if session("ModMain") = "WK" and ( session("MODSTART") <> "XP" or (session("MODSTART") = "XP" and session("WKFCODIGO") = sCodigo) ) then
				Solapas(0)  = "WORKFLOW"
				Solapas2(0) = ASPs(0)

				''' xstrsol     = 0
				xstrsol     = 0
				If session("APLMDCLI") <> "BC" Then
					' se saco porque generaba error con la solapa propiedades en el primer lugar 22_9_2017
					'  if sSClick      = 99 then sSClick      = solord(1)
				End If
			Else
				xstrsol     = 1
			End If





			X=0

			If sFuncion <> "E"  Then

				'if sFuncion <> "E" then

				Solapas(20)  = "PORTADA"
				Solapas2(20) = ASPs(20)


				Strsql = "SELECT OBJportada.detcon, OBJportada.detsecuen, OBJportada.dettipo from OBJportada "
				Strsql = Strsql & " where OBJportada.codobjeto=" & sCodigo & " and OBJportada.detcon > 0 and OBJportada.dettipo <> 'N' "

				' response.write strsql
				Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsxx1.eof Then


					For xiord=0 To 19

						If solord(xiord) = 1 Then

							xmovey = 20

							'  for xiord2=xiord + 1 to 19
							'      xmoven           = solord(xiord2)
							'      solord(xiord2)   = xmovey
							'      xmovey           = xmoven
							'  next



							Exit For
						End If
					Next


					WWWMOVER = 20

					For xiord=1 To 20

						XXXMOVER      = solord(xiord)
						solord(xiord) = WWWMOVER
						WWWMOVER      = XXXMOVER

					Next

				End If


			End If

			For z = 2 To 16

				If rs(z) Then




					If (sFuncion = "E"  ) Or (sFuncion = "C" And z <> 2 And z <> 4 And z <> 5 ) Or (sFuncion = "CD" And z <> 2 And z <> 5 )  Or (sFuncion = "D"  And z = 4 ) Then
						x = x + 1
						Solapas(z-1)  = Replace(RS(z + 15)," ","*")
						Solapas2(z-1) = ASPs(z-1)


						If sFuncion <> "E" And z = 6 Then

							Solapas(16)  = "CIRCUITO COMPLETO"
							Solapas2(16) = ASPs(16)

							For xiord=0 To 19

								If solord(xiord) = 5 Then

									xmovey = 16

									For xiord2=xiord + 1 To 19
										xmoven           = solord(xiord2)
										solord(xiord2)   = xmovey
										xmovey           = xmoven
									Next



									Exit For
								End If
							Next


						End If

					End If
				End If

			Next


			xfinal = UBound(Solapas)

		End If

		''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

		Call ofv.cerrarconsulta(rs)


	End If


	TABELEN = ""
	FTABELEN = ""



	xSolasp = ""
	xSolasp = xSolasp & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
	xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
	xSolasp = xSolasp & "&Funcion=" & sFuncion
	xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
	xSolasp = xSolasp & "&SSCLICK="


	xSolaspd = ""
	xSolaspd = xSolaspd & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
	xSolaspd = xSolaspd & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
	xSolaspd = xSolaspd & "&Funcion=" & sFuncion
	xSolaspd = xSolaspd & "&SSCLICK="

	r1 = ""


	If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
		xGenConta = ofv.GetTit(CLng(ofv.ModTitCls) + 5)
		xGenConta = ofv.FTitDat1

		xsclase = "xts"

	Else



		xsclase = "xts"

	End If

	xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
	xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"



	xsolini = 99


	If session("WEBDESK") = "S" And session("ModMain") = "WK" Then

		sInput  = "<font color='#990000' >"
		sInput  = sInput & "Este documento requiere su confirmacion en el workflow!!!</font>"
		sLink   = sInput

		CTABELEN = ofv.GenCelda("","vT  xstyle='font-size:8pt;background-color:#ffffff;border-style:solid;border-width:1;border-color:#000000;padding-left:5px;' ","","","","","",sLink,"si")
		'  FTABELEN = FTABELEN & ofv.GenRow("","","","","25","","",CTABELEN,"si")

	End If

	If session("WEBDESK") = "S" And session("ModMain") <> "WK" And (sFuncion = "C" Or sFuncion = "CD") Then

		xverasp = "CoObjRecomendar.asp?Tipobjeto=" & sTipo & "&Codobjeto="
		xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
		xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5


		sInput  = "<b><font color='#ff0000' size=2  >&#058;&#0183;)</font></b>&nbsp;&nbsp;"
		'''  sInput  = ""
		sInput  = sInput & "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >Recomendar</font>"
		'''  sInput  = sInput & "&nbsp;<font color='#ff0000' size=2 >&#033;</font>"
		sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:bold;'  ", xverasp, "Principal", "", "", sInput, "si")
		Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

		wwwestilo = " xstyle='background-color:#FFEEDD;border-style:dotted;border-width:2;border-color:#990000;padding-left:10px;' "

		CTABELEN = ofv.GenCelda("","uTbrr" & wwwestilo & " ","","","","","",sLink,"si")
		' FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")

	End If


	For x = xstrsol To xfinal
		'' for x = xstrsol to 2

		If solapas(solord(x)) <> "N" Then
			sSolapa =  solapas(solord(x))

			If  xsolini = 99 Then  xsolini = solord(x)

			If  sSClick = 99 Then  sSClick = xsolini

			If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


				sInput  = Replace(UCase(sSolapa),"*"," ")
				xsclase = "at xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:10px;'"


			Else

				sInput  = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
				xsclase = "at xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:10px;'"

			End If

			If session("WEBDESK") = "S" Then

				sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >" & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;") & "</font>"

				'   ssMnnSOL = "MNNSOL" & X
				iitemC=iitemC+1
				'   sLink   = "<span id='" & ssMnnSOL & "'  class='XX' >" & sInput & "</span> "
				'   sLink   = OFV.GenLink("", "cmb  onclick=" & chr(34) & "SetMnnSOL('"&ssMnnSOL&"');" & chr(34) & "style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")
				sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput , "si")
				'
				Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				wwwestilo = " xstyle='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

				If (solord(x)) = 0 Then
					wwwestilo = " xstyle='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
				End If


				CTABELEN = ofv.GenCelda("","utbrr" & wwwestilo & " ","","","","","",sLink,"si")
				FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")



				If sFuncion = "E" Or sFuncion = "ED"  Then

				Else
					xxxx = GetIndice(solord(x))
				End If


			Else

				t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
				t3 = t3 & sInput
				t3 = t3 & "</A>"

				t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

				r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

			End If


			If CLng(sSClick) = CLng(solord(x)) Then
				session("codSol")=sSClick
				' response.write clng(sSClick)
				xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
				xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
				xverasp = xverasp & "&GLBSOLN=S" & solord(x)
				xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 '& "&ddd=" & sFuncion

				sHTML1 = ""

				sHTML1 = sHTML1 & "<form name=SSOLACT id=SSOLACT  action='" & xverasp & "' "
				sHTML1 = sHTML1 & " method=post target='"
				If CLng(sSClick) = 12 Then
					sHTML1 = sHTML1 & "_blank"
				Else
					sHTML1 = sHTML1 & "Principal"
				End If
				sHTML1 = sHTML1 & "' style='display:none;' >" & Chr(13)
				sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & Chr(13)
				sHTML1 = sHTML1 & "</form>" & Chr(13)

				'''''''''''''''''''''''''''''''''
				' NUEVO PARA VERSION 4.6

				'response.write request.querystring("Test") & " pregunta por Test " & "<br>"

				If request.querystring("Test") = "T" Then

					sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
					sHTML1 = sHTML1 & " window.top.CorSol='" & Replace(UCase(sSolapa),"*"," ") & "'; " & Chr(13)
					sHTML1 = sHTML1 & "</script>" & Chr(13)

				Else
					sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
					sHTML1 = sHTML1 & " window.top.CorSol='" & Replace(UCase(sSolapa),"*"," ") & "'; " & Chr(13)

					sHTML1 = sHTML1 & " var element =  document.getElementById(window.top.imenu);"

					sHTML1 = sHTML1 & " if (typeof(element) != 'undefined' && element != null)"
					sHTML1 = sHTML1 & " {"

					''''''''Agregado 10032015 - Version 4.6
					sHTML1 = sHTML1 & " element.style.backgroundColor='"&session("CMICTT")&"';"
					sHTML1 = sHTML1 & " element.style.color='"&session("CMICTTFT")&"';"
					'sHTML1 = sHTML1 & "var f=element;for (i = 0; i < 5; i++) {"
                   '  sHTML1 = sHTML1 & "var f=f.parentElement;	"
					'sHTML1 = sHTML1 & " if (f.tagName=='TD'){break;}};"
                  '  sHTML1 = sHTML1 & " f.innerHTML='<div class=""item-container active-item""> ' + f.innerHTML+'</div>' ;"
					''''''''Agregado 10032015 - FIN - Version 4.6

					sHTML1 = sHTML1 & " }"

					sHTML1 = sHTML1 & "</script>" & Chr(13)

				End If
				' NUEVO PARA VERSION 4.6 -FIN-
				'''''''''''''''''''''''''''''''''

				If session("WEBDESK") = "S" Then




				Else

					If  Session("FULLGUI") = "N" Then
						styl = "background-color:#888888;"
					Else
						styl = XspEx("#888888","#cccccc",0) & ";"
					End If

					xlfun = " onclick=""document.getElementById('MDBEDL').style.display=''; document.getElementById('DBEDTD1').style.display='none'; document.getElementById('DBEDTD2').style.display=''; document.getElementById('MM2').style.display=''; "" "
					xlfun2 = " onclick=""document.getElementById('MDBEDL').style.display='none'; document.getElementById('DBEDTD1').style.display=''; document.getElementById('DBEDTD2').style.display='none'; document.getElementById('MM2').style.display='none'; "" "

					macel = ofv.GenCelda("","at xstyle='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:10px;' " & xlfun,"","90%","","","",Replace(UCase(sSolapa),"*"," "),"si")

					xsclase2 = "at xstyle='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


					macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
					macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

					marow = ofv.GenRow("","","","",lengtd,"","",macel,"si")

					sHTML = sHTML & OFV.GenTabla("DBED", "aaa xstyle='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

				End If

			End If


		End If
	Next

	If SMAINAPL = "OR" Or SMAINAPL = "PL" Then

	Else

		''-----------------------------------------------------------------------------------------------------

		''''''''''''''''''''''''''''''''''''''''''''
		'' Accesos directos del documento ''''''''''
		''''''''''''''''''''''''''''''''''''''''''''

		Dim ATABSCPT(200,2)

		ISCPT = 0
		TSCPT = 0

		''    response.write "entro1"

		If sFuncion <> "X"   Then

			''  response.write "entro2"

			ReDim Preserve bConsultar(4)



				bConsultar(2) = 0
				bConsultar(4) = 0

				bConsultar(1) = 1
				bConsultar(3) = 1





				bConsultar(2) = 0
				bConsultar(4) = 0

				bConsultar(1) = 0
				bConsultar(3) = 1



			sVerEstado = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

			sverestado = sverestado & " and Objprty.Estado In (9,"
			For i = 1 To 4
				If bConsultar(i) Then sVerEstado = sVerEstado & i & ","
			Next
			sVerEstado = Left(sVerEstado, Len(sVerEstado) - 1) & ") "

			'''''''''''''''''''''''''''''''''''''''''



			sActcon  = "Consultar.asp?Codigo=NNNN&Tipo=TTTT"
			sActcon  = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
			sHref    = "&nbsp;"
			' response.write sActcon
			sFuncion = " onclick=" & Chr(34) & " window.top.abreWin('" & sActcon & "','TTTT'); return false;" & Chr(34)


			Strsql = "SELECT * FROM objacd  "
			Strsql = Strsql & " Where codobjeto=" & sCodigo
			Strsql = Strsql & " And tipobjeto = '" & stipo & "' "

			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do Until rs.eof

				''   response.write "entro3"


				If  rs(3) = "S" Then
					'''' Carpeta superior

					If  rs(6) = "S" Then

						''----------------

						strsql  = "select * from csbarr where codojasoc = " & sCodigo
						Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
						Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rs.eof Then
							essCodigo = rs1(1)
							essTipo   = rs1(2)

						Else
							essCodigo = 0
							essTipo   = "NO"
						End If
						Call ofv.cerrarconsulta(rs1)


						sNIVE   = "no"

						StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
						StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
						Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
						If Not rsIcono.EOF Then

							sIcono  = rsIcono(0)
							sTipoE  = rsIcono(1)
							sPerfil = rsIcono(2)
							sNIVE   = rsIcono(3)

						End If
						Call OFV.cerrarconsulta(rsIcono)

						If sNIVE   = "C"  Then

							strsql  = "select * from csbarr where codojasoc = " & essCodigo
							Strsql  = Strsql & " and tipobjasoc = '" & essTipo & "'"
							Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If  Not rs.eof Then

								essCodigo = rs1(1)
								essTipo   = rs1(2)

							End If
							Call ofv.cerrarconsulta(rs1)


							'''''''''''''''''''''''''''''''''''''''''

							StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
							StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
							StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
							StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
							StrSql = StrSql & " where "
							StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
							StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
							StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
							StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
							StrSql = StrSql & "  objetos.objun =  1 )  and "
							StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
							StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
							StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
							''  response.write "<br>" & strsql & "<br>"
							sVino = "N"
							Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If  Not rs2.EOF Then

								Do Until rs2.EOF

									If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

										sVinook = "N"
										If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
											'' response.write "<br><br>entro"
											sVino   = "S"
											sVinook = "S"
										End If

										If SESSION("EMPMUN") = "S" Then
											StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
											StrSql = StrSql & " where "
											StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
											StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
											Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If  Not rsmas2.EOF Then
												sVino   = "S"
												sVinook = "S"
											End If
										End If

										sNIVE   = "D"

										If  sVinook = "S" Then

											stitem = rs2(4)
											StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
											StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
											Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If Not rsIcono.EOF Then
												sIcono  = rsIcono(0)
												sTipoE  = rsIcono(1)
												sPerfil = rsIcono(2)
												sNIVE   = rsIcono(3)

												sitem = sTipoE & ": " & rs2(1)

											End If
											Call OFV.cerrarconsulta(rsIcono)

											If sNIVE   = "D"  Then


												sSolapa =  sitem

												If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

													sInput = Replace(UCase(sSolapa),"*"," ")

												Else

													sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"

												End If

												sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
												sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))





												If session("WEBDESK") = "S" Then


													sInput  = ">" & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;") & ""
													sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
													Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

													wwwestilo = " xstyle='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

													CTABELEN = ofv.GenCelda("","utbrr" & wwwestilo & " ","","","","","",sLink,"si")
													FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")




												Else

													t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
													t3 = t3 & sInput
													t3 = t3 & "</A>"

													t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

													r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

												End If





											Else





											End If


										End If

									End If

									rs2.movenext
								Loop


							End If



							'''''''''''''''''''''''''''''''''''''''

						End If

						''----------------



					End If





				ElseIf  rs(3) = "P" Then
					'''' Carpeta primaria

					If  rs(6) = "S" Then

						strsql  = "select * from csbarr where codojasoc = " & sCodigo
						Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
						Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rs.eof Then
							essCodigo = rs1(1)
							essTipo   = rs1(2)

						Else
							essCodigo = 0
							essTipo   = "NO"
						End If
						Call ofv.cerrarconsulta(rs1)


						sNIVE   = "no"

						StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
						StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
						Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
						If Not rsIcono.EOF Then

							sIcono  = rsIcono(0)
							sTipoE  = rsIcono(1)
							sPerfil = rsIcono(2)
							sNIVE   = rsIcono(3)

						End If
						Call OFV.cerrarconsulta(rsIcono)




						StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
						StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
						StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
						StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
						StrSql = StrSql & " where "
						StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
						StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
						StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
						StrSql = StrSql & " (objetos.objun =  0" & session("UN") & " or "
						StrSql = StrSql & "  objetos.objun =  1 )  and "
						StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
						StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
						StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
						''   response.write "<br>" & strsql & "<br>"
						sVino = "N"
						Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
						If  Not rs2.EOF Then

							Do Until rs2.EOF

								If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

									sVinook = "N"
									If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
										'' response.write "<br><br>entro"
										sVino   = "S"
										sVinook = "S"
									End If

									If SESSION("EMPMUN") = "S" Then
										StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
										StrSql = StrSql & " where "
										StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
										StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
										Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
										If  Not rsmas2.EOF Then
											sVino   = "S"
											sVinook = "S"
										End If
									End If

									sNIVE   = "D"

									If  sVinook = "S" Then

										stitem = rs2(4)
										StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
										StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
										Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
										If Not rsIcono.EOF Then
											sIcono  = rsIcono(0)
											sTipoE  = rsIcono(1)
											sPerfil = rsIcono(2)
											sNIVE   = rsIcono(3)

											sitem = sTipoE & ": " & rs2(1)

										End If
										Call OFV.cerrarconsulta(rsIcono)

										If sNIVE   = "D"  Then


											sSolapa =  sitem

											If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
												sInput = Replace(UCase(sSolapa),"*"," ")
											Else
												sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
											End If

											sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
											sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))




											If session("WEBDESK") = "S" Then




												'    sInput  = "" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
												'    sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:normal;'  " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
												'    Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

												'    CTABELEN = ofv.GenCelda("","UT  style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' ","","","","","",sLink,"si")
												'    FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



												sInput  = ">" & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;") & ""
												sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
												Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

												wwwestilo = " xstyle='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

												CTABELEN = ofv.GenCelda("","utbrr" & wwwestilo & " ","","","","","",sLink,"si")
												FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")




											Else

												t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
												t3 = t3 & sInput
												t3 = t3 & "</A>"

												t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

												r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

											End If


										Else


										End If


									End If

								End If

								rs2.movenext
							Loop


							'''''''''''''''''''''''''''''''''''''''

						End If




					End If




				ElseIf  rs(3) = "R" Then
					'''' Carpeta dependiente

					If  rs(6) = "S" Then

						''----------------

						TSCPT = TSCPT + 1

						ATABSCPT(TSCPT,0) = ""
						ATABSCPT(TSCPT,1) = sCodigo
						ATABSCPT(TSCPT,2) = sTipo



						strsql  = "select * from csbarr where codojasoc = " & sCodigo
						Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
						Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If  Not rs.eof Then
							essCodigo = rs1(1)
							essTipo   = rs1(2)

						Else
							essCodigo = 0
							essTipo   = "NO"
						End If
						Call ofv.cerrarconsulta(rs1)


						sNIVE   = "no"

						StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
						StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
						Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
						If Not rsIcono.EOF Then

							sIcono  = rsIcono(0)
							sTipoE  = rsIcono(1)
							sPerfil = rsIcono(2)
							sNIVE   = rsIcono(3)

						End If
						Call OFV.cerrarconsulta(rsIcono)

						If sNIVE   = "C"  Then

							'  essCodigo = sCodigo
							'  essTipo   = sTipo

							'''''''''''''''''''''''''''''''''''''''''

							StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
							StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
							StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
							StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
							StrSql = StrSql & " where "
							StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
							StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
							StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
							StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
							StrSql = StrSql & "  objetos.objun =  1 )  and "
							StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
							StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
							StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
							''   response.write "<br>" & strsql & "<br>"
							sVino = "N"
							Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
							If  Not rs2.EOF Then

								Do Until rs2.EOF

									If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then

										sVinook = "N"
										If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
											'' response.write "<br><br>entro"
											sVino   = "S"
											sVinook = "S"
										End If

										If SESSION("EMPMUN") = "S" Then
											StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
											StrSql = StrSql & " where "
											StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
											StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
											Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If  Not rsmas2.EOF Then
												sVino   = "S"
												sVinook = "S"
											End If
										End If

										sNIVE   = "D"

										If  sVinook = "S" Then

											stitem = rs2(4)
											StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
											StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
											Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
											If Not rsIcono.EOF Then
												sIcono  = rsIcono(0)
												sTipoE  = rsIcono(1)
												sPerfil = rsIcono(2)
												sNIVE   = rsIcono(3)

												sitem = sTipoE & ": " & rs2(1)

											End If
											Call OFV.cerrarconsulta(rsIcono)

											If sNIVE   = "D"  Then




											Else


												TSCPT = TSCPT + 1

												ATABSCPT(TSCPT,0) = ""
												ATABSCPT(TSCPT,1) = rs2(0)
												ATABSCPT(TSCPT,2) = rs2(2)


											End If


										End If

									End If

									rs2.movenext
								Loop

							End If



							Do While Not ISCPT > TSCPT

								If ISCPT = 0 Then ISCPT = 1


								For ISCPT=ISCPT To TSCPT



									StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
									StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
									StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
									StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
									StrSql = StrSql & " where "
									StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
									StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
									StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
									StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
									StrSql = StrSql & "  objetos.objun =  1 )  and "
									StrSql = StrSql & " RELAOBJ.codobjmaster = 0" & ATABSCPT(ISCPT,1)
									StrSql = StrSql & " and relaobj.tipobjmaster = '" & ATABSCPT(ISCPT,2) & "' " & sVerEstado
									StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
									''    response.write "<br>" & strsql & "<br>"
									sVino = "N"
									Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
									If  Not rs2.EOF Then

										Do Until rs2.EOF

											If CStr(rs2(0) & "")  <> CStr(scodigo & "") Then



												sVinook = "N"
												If  CStr(rs2(12) & "") = CStr(session("UN") & "") Then
													'' response.write "<br><br>entro"
													sVino   = "S"
													sVinook = "S"
												End If

												If SESSION("EMPMUN") = "S" Then
													StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
													StrSql = StrSql & " where "
													StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
													StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
													Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
													If  Not rsmas2.EOF Then
														sVino   = "S"
														sVinook = "S"
													End If
												End If

												sNIVE   = "D"

												If  sVinook = "S" Then

													stitem = rs2(4)
													StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
													StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
													Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
													If Not rsIcono.EOF Then
														sIcono  = rsIcono(0)
														sTipoE  = rsIcono(1)
														sPerfil = rsIcono(2)
														sNIVE   = rsIcono(3)

														sitem = sTipoE & ": " & rs2(1)

													End If
													Call OFV.cerrarconsulta(rsIcono)

													If sNIVE   = "D" Then


														sSolapa =  sitem

														If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
															sInput = Replace(UCase(sSolapa),"*"," ")
														Else
															sInput = "&nbsp;" & Replace(UCase(sSolapa),"*","&nbsp;") & "&nbsp;"
														End If

														sFuncionDSP = Replace(sFuncion,"NNNN",rs2(0))
														sFuncionDSP = Replace(sFuncionDSP,"TTTT",rs2(2))



														If session("WEBDESK") = "S" Then



															'  sInput  = "" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & ""
															'  sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:normal;'  " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
															'  Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

															'  CTABELEN = ofv.GenCelda("","UT  style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' ","","","","","",sLink,"si")
															'  FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")


															sInput  = ">" & Replace(Capitalize(Replace(UCase(sSolapa),"*"," "))," ","&nbsp;") & ""
															sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;' " & sFuncionDSP & " ", "javascript:void(0);", "", "", "", sInput, "si")
															Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

															wwwestilo = " xstyle='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

															CTABELEN = ofv.GenCelda("","utbrr" & wwwestilo & " ","","","","","",sLink,"si")
															FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")


														Else

															t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
															t3 = t3 & sInput
															t3 = t3 & "</A>"

															t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

															r1 = r1 & ofv.GenRow("","","","",lengtd,"","",t4,"si")

														End If



													Else



														TSCPT = TSCPT + 1

														ATABSCPT(TSCPT,0) = ""
														ATABSCPT(TSCPT,1) = rs2(0)
														ATABSCPT(TSCPT,2) = rs2(2)


													End If


												End If


											End If

											rs2.movenext
										Loop

									End If


								Next

							Loop



						End If



					End If


				Else

					''''' documentos o subcarpetas independientes


				End If


				''''''''''''''''''''''''''''


				'''''''''''''''''''''''''''''



				rs.movenext
			Loop

		End If

		''''''''''''''''''''''''''''''''''''''''''''


		''-----------------------------------------------------------------------------------------------------


	End If



	If session("WEBDESK") = "S" Then


		'''    sHTML = sHTML & Celda
		'''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

		sHTML = sHTML & "<div id='WKTBSP1'  style='display:none'>"
		sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "'aaa tnav' ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN, "si")

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
			sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa xstyle='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
		Else
			sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa xstyle='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
		End If

		sHTML = sHTML & "</div>"

	End If








	sHTML = sHTML & sHTML1
    sHTML = sHTML & armasolapa(sCodigo, sTipo,sFuncionSol)
	Call ofv.cerrarconn(cn)




	sHTML = sHTML & "</div></body>"
	sHTML = sHTML &  "</html>"







End If
If session("APLMDCLI")="BGAL" Then
	Response.Write Replace(sHTML,"&nbsp;"," ") 'sHTML
Else
	Response.Write SHTML
End If

Function GetIndice(gtsolapa)



	wwsNIV1     = gtsolapa
	wwsNIV2     = "0"
	wwsNIV3     = "0"
	wwsNIV4     = "0"
	wwsNIV5     = "0"


	sft   = "<font face=tahoma,arial size=1 color=#336699>&nbsp;<b>"
	sfta  = "<font face=tahoma,arial size=1 color=#993333>&nbsp;<b>"
	sfidc = "<font face=corflaticon size=2 color=#336699>0</font>"
	sfida = "<font face=corflaticon size=2 color=#993333>1</font>"
	sfitc = "<font face=corflaticon size=2 color=#336699>0</font>"
	sfita = "<font face=corflaticon size=2 color=#993333>1</font>"
	sftc  = "</b></font>"


	'  ASPs(1)  = "edPropiedades.asp?SKIND=1"
	'  ASPs(2)  = "CONWEDT.asp?SKIND=1"
	'  ASPs(3)  = "CoSectores.asp?SKIND=1"
	'  ASPs(4)  = "edEventos.asp?SKIND=1"
	'  ASPs(5)  = "CoProcedimientos.asp?SKIND=1"
	'  ASPs(6)  = "CoObjDsc.asp?SKIND=1"
	'  ASPs(7)  = "CoControles.asp?SKIND=1"
	'  ASPs(8)  = "CoEsquemas.asp?SKIND=1"
	'  ASPs(9)  = "CoTablasCons.asp?SKIND=1"
	'  ASPs(10) = "CoArchCons.asp?SKIND=1"
	'  ASPs(11) = "CoestrCons.asp?SKIND=1"
	'  ASPs(12) = "CursogramaLnch.asp?SKIND=1"
	'  ASPs(13) = "CoArchCons.asp?SKIND=2"
	'  ASPs(14) = "CoFAQ.asp?SKIND=1"
	'  ASPs(15) = "CoGeneral.asp?SKIND=1"
	'  ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"

	gtselinx = ""


	Select Case CStr(gtsolapa & "")

		Case "1"

		Case "2"

		Case "3"

		Case "4"

		Case "5"

		eveant = ""

		Strsql = "SELECT *  "
		Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & sCodigo
		''  Strsql = Strsql & sSelecx
		'' strsql = strsql & addperfil
		strsql = strsql & " order by secuen, secuencia,codproced "
		''response.write strsql
		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		Do While Not rsx.eof

			If eveant <> rsx(4) Then

				'    Celdat  = "<td class=uts >" & sfidkA & "</td>"
				'    Celdat  = Celdat & "<td class=utg  >" & rsx(2) & ".&nbsp;" & rsx(4) & "</td>"
				'    Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

				eveant  = rsx(4)
			End If

			sAccion = "CoProcedimientos.asp?INXMOD=N&CantRegAMover=A&Codobjeto=" & sCodigo & "&Menu=" & sMenux
			sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&proced=" & rsx(0) & "&test=" & sTest
			sOnClik1 = ""
			sOnClik2 = " onclick=""document.getElementById('ASC" & rsx(0) & "').click();"" "

			'  if rsx(0)  = rs(0).value then
			'     sfidkD  = sfidkc
			'     stexto1 = "<u>" & rsx(6) & ".&nbsp;" & rsx(3) & "&nbsp;-&nbsp;" & rsx(5) & "</u>"

			'  else
			sfidkD  = sfidkB
			stexto1 = rsx(2) & "." & rsx(6) & "."  & Capitalize(rsx(4) & "")  & "." & Capitalize(rsx(3) & "") & " - " & Capitalize(rsx(5) & "")
			stexto1 = Replace(stexto1," ","&nbsp;")

			'  end if

			'  Celdat  = stexto1

			'  stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=nt href='" & sAccion & "' " & sOnClik1 & " "
			'  stexto1 = stexto1 & "   >" & Celdat & "</a>"

			'  Celdat  = "<td class=utis >" & sfidkD & "</td>"
			'  Celdat  = Celdat & "<td class=uti style='cursor:pointer;' " & sOnClik2 & " >" & stexto1 & "</td>"
			'  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

			wwsNIV2     = (rsx(0) & "")
			wwsNIV3     = "0"

			sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

			sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
			iitemC=iitemC +1

			sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput , "si")
			Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp</span> "


			CTABELEN = ofv.GenCelda("","utbrr xstyle='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink ,"si")
			FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")




			rsx.movenext
		Loop



		Case "6"

		Case "7" '' controles


		wwbbusca = "0"

		Strsql = "SELECT * from cscontrolsel "
		Strsql = Strsql & " Where codobjeto = " & sCodigo
		Strsql = Strsql & " and seldes is NOT NULL and descri is NOT NULL "
		'''''''Strsql = Strsql & " and descri is NOT NULL "

		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		'contar = 0
		Do While Not rsx.eof

			'contar = contar + 1
			'response.write contar & "contar " & "<br>"

			If CStr(rsx(0) & "")   = wwbbusca Then
				sfidkD  = sfidkC
				stexto1 = rsx(1)
				stextoact = stexto1
			Else
				sfidkD   = sfidkA
				stexto1  = rsx(1) & ""
				''
			End If
			''
			sAccion  = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
			sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & rsx(0) & "&test=" & sTest
			sOnClik1 = ""
			sOnClik2 = " onclick=""document.getElementById('ASC" & rsx(0) & "').click();"" "

			'    Celdat  = stexto1
			'    stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
			'    stexto1 = stexto1 & "   >" & Celdat & "</a>"

			'     Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
			'     Celdat  = Celdat & "<td class=ut  >" & stexto1  & "zz" & "</td>"
			'     Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

			stexto1 = Replace(Capitalize(stexto1 & "")," ","&nbsp;")

			wwsNIV2     = CStr(rsx(0) & "")
			wwsNIV3     = "0"

			sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

			sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
			iitemC=iitemC +1
			sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
			Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

			CTABELEN = ofv.GenCelda("","utbrr xstyle='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
			FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")


			''
			'end if
			''

			rsx.movenext
		Loop



		Case "8" '' esquemas

		wwbbusca = "0"

		Strsql = "SELECT * from csesquemasel "
		Strsql = Strsql & " Where codobjeto = " & sCodigo

		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		Do While Not rsx.eof

			If CStr(rsx(0) & "")   = wwbbusca Then
				sfidkD  = sfidkC
				stexto1 = rsx(1)
				stextoact = stexto1
			Else
				sfidkD   = sfidkA
				stexto1  = rsx(1) & ""
				sAccion  = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & sCodigo & "&Menu=" & sMenux
				sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & rsx(0) & "&test=" & sTest
				sOnClik1 = ""
				sOnClik2 = " onclick=""document.getElementById(ASC" & rsx(0) & "').click();"" "

				'    Celdat  = stexto1
				'    stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
				'    stexto1 = stexto1 & "   >" & Celdat & "</a>"

				'     Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
				'     Celdat  = Celdat & "<td class=ut  >" & stexto1  & "zz" & "</td>"
				'     Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

				stexto1 = Replace(Capitalize(stexto1 & "")," ","&nbsp;")

				wwsNIV2     = CStr(rsx(0) & "")
				wwsNIV3     = "0"

				sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

				sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
				iitemC=iitemC+1
				sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
				Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				CTABELEN = ofv.GenCelda("","utbrr xstyle='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
				FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

			End If

			rsx.movenext
		Loop



		Case "9"  ''' bloque de tablas

		Strsql = "SELECT * "
		Strsql = Strsql & " from ObjTabDef Where codobjeto = " & sCodigo
		Strsql = Strsql & " order by tabsecuen "

		Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsxx1.eof Then
			'' if sdocum = "" then sdocum = cstr(rsxx1(0))

			' tomo todas las tablas

			stexto1 = "Seleccionar&nbsp;tabla"
			'  Celdat  = stexto1

			sAccion = "javascript:void(0);"
			sOnClik = " onclick=""document.getElementById('TIND2').style.display='';document.getElementById('TIND1').style.display='none';"" "


			stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
			stexto1 = stexto1 & "   >" & Celdat & "</a>"


			'  Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
			'  Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"


			stexto1 = "Tablas"

			'  Celdat  = "<td class=uts >" & sfidk0 & "</td>"
			' Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
			' Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")


			Do Until rsxx1.eof

				wwsarch  = "0" & rsxx1(0)



				'    if  sdocum = cstr(rsxx1(0)) then
				'        sarch  = rsxx1(0)
				'    end if

				'   if  rsxx1(3) <> "" then
				''  Celda  = ofv.GenCelda("","cc","","1%","","","",stexto1,"si")
				''  Celda  = Celda & ofv.GenCelda("","","","2","","","","&nbsp;","si")
				''  Formulario = Formulario & Celda
				'  end if


				'   if  sdocum = cstr(rsxx1(0)) then
				'       sfidkD  = sfidkC
				'       stexto1 = replace(rsxx1(4)," ","&nbsp;")
				'       stextoact = stexto1

				'   else
				sfidkD  = sfidkA
				''''    stexto1 = replace(rsxx1(4)," ","&nbsp;")

				stexto1 = Replace(Capitalize(rsxx1(4) & "")," ","&nbsp;")
				sAccion = "CoTablasCons.asp?INXMOD=N&Codobjeto=" & scodigo & "&Tipobjeto=" & stipo & "&sdocum=" & rsxx1(0)
				sOnClik1 = ""
				sOnClik2 = " onclick=""document.getElementById('ASC" & rsxx1(0) & "').click();"" "


				Celdat  = stexto1

				'  stexto1 = "<A ID='ASC" & rsxx1(0) & "' NAME='ASC" & rsxx1(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
				'  stexto1 = stexto1 & "   >" & Celdat & "</a>"

				'  Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
				'  Celdat  = Celdat & "<td class=uti  >" & stexto1 & "</td>"
				'  Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

				wwsNIV2     = (rsxx1(0) & "")
				wwsNIV3     = "0"

				sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

				sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & "</font>"
				iitemC=iitemC+1
				sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
				Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

				CTABELEN = ofv.GenCelda("","utbrr xstyle='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:25px;' ","","","","","",sLink,"si")
				FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")


				'  end if

				Strsql = "SELECT * From Objtabblk Where Tabdef = " & wwsarch & " order by tabblksec"
				Set rsxx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				Do Until rsxx2.eof
					wwnID      = rsxx2(0)
					wwnSec     = rsxx2(4)
					wwnBTitulo = rsxx2(5)

					If Len(wwnBTitulo) > 0 Then
						wwltitu = Len(wwnBTitulo)
						If  wwnBTitulo <> Space(wwltitu) And UCase(wwnBTitulo) <> "NO" Then

							stexto1 = wwnSec & "." & Replace(Capitalize(wwnBTitulo & "")," ","&nbsp;")
							stexto1 = Replace(stexto1," ","&nbsp;")

							'  Celdat  = stexto1
							'  stexto1 = "<A class=ut href='javascript:void(0);' "
							'  stexto1 = stexto1 & "   >" & Celdat & "</a>"
							'  Celdat  = "<td class=utis style='cursor:pointer;' onclick=""document.all['TBLMAIN'].style.visibility='hidden';creawin('" & sPoint & "');document.all['TBLMAIN'].style.visibility='visible';"" >" & sfidkB & "</td>"
							'  Celdat  = Celdat & "<td class=uti onclick=""document.all['TBLMAIN'].style.visibility='hidden';creawin('" & sPoint & "');document.all['TBLMAIN'].style.visibility='visible';"" >" & stexto1 & "</td>"
							'  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

							wwsNIV2     = (rsxx1(0) & "")
							wwsNIV3     = (wwnID & "")

							sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

							sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
							iitemC=iitemC+1
							sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
							Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;&nbsp;&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

							CTABELEN = ofv.GenCelda("","utbrr  xstyle='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:40px;' ","","","","","",sLink,"si")
							FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

						End If
					End If

					rsxx2.movenext
				Loop


				rsxx1.movenext
			Loop

		Else


		End If


		Case "10"

		Case "11" '' bloque de contenido ----------------------------------------------------------------------------------------

		Dim Inx
		Dim nrs(100)

		ReDim Preserve bsec(8)
		ReDim Preserve bnum(8)
		ReDim Preserve bnumG(8)


		Dim Pdef(13,18)

		strsql = "Select * From Objestr Where codobjeto =" & sCodigo
		Strsql = Strsql & " order by CONsecuen "
		'  response.write strsql & "<br>"
		Set rsxx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		bsecuen = ""
		bDescri = " "
		bobj = ""
		Do While Not rsxx1.eof
			bsecuen = rsxx1(4)
			bDescri = rsxx1(3)
			bobj    = rsxx1(0)

			'  stexto1 = "<A ID='ATIT" & rsxx1(0) & "' NAME='ATIT" & rsxx1(0) & "' "  & chr(13)
			'  stexto1 = stexto1 & " class=nt href='coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
			'  stexto1 = stexto1 & "&sdocum=" & sdocum & "&sdocc=" & rsxx1(0)
			'  stexto1 = stexto1 & "'   >" & sfidkc & "&nbsp;<big>" & rsxx1(3) & "</big></a>" & chr(13)

			'  Celdat = "<td class=ut style='cursor:pointer;'  onclick=""document.all.ATIT" & rsxx1(0) & ".click();this.style.backgroundColor='red';"" >aaaaaaa" & stexto1 & "qqqqqq</td>"
			'  Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si") & chr(13)

			stexto1 = bsecuen & "." & Replace(Capitalize(bDescri & "")," ","&nbsp;")
			stexto1 = Replace(stexto1," ","&nbsp;")

			wwsNIV2     = bobj & ""
			wwsNIV3     = "0"

			sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

			sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;' >" & stexto1 & " </font>"
			iitemC=iitemC+1
			sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
			Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

			CTABELEN = ofv.GenCelda("","utbrr xstyle='background-color:#E9E9E9;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#ffaa22;padding-left:10px;' ","","","","","",sLink,"si")
			FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")

			wentro = 0

			strsql = "Select * From SysEstprty"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			t = 0
			Do Until rs.eof
				For ti = 1 To 18
					If (rs(ti) & "") = "" Then
						sti = ""
					Else
						sti = rs(ti)
					End If

					Pdef(t,ti - 1) = sti

				Next
				rs.movenext
				t = t + 1
			Loop

			Call ofv.cerrarconsulta(rs)

			strsql = "Select * From ObjEstPrty where detcon = " & bobj
			'response.write strsql & "<br>"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			t = 0
			Do Until rs.eof
				Select Case rs(4)
					Case "A"
					t = 0
					Case "B"
					t = 1
					Case "C"
					t = 2
					Case "D"
					t = 3
					Case "E"
					t = 4
					Case "F"
					t = 5
					Case "G"
					t = 6
				End Select

				For ti = 5 To 22
					If (rs(ti) & "") = "" Then
						sti = ""
					Else
						sti = rs(ti)
						Pdef(t,ti - 5) = sti
					End If



				Next
				rs.movenext
			Loop

			Call ofv.cerrarconsulta(rs)

			PInden = Pdef(0,7)
			If PInden <> "NO" Then PInden = "SI"

			PNumer = Pdef(0,12)
			If PNumer <> "NO" Then PNumer = "SI"

			PNumCls = Pdef(0,14)
			ReDim Preserve PNumTip(13)
			For i=0 To 13
				PNumTip(i) = Pdef(i,13)
				If PNumTip(i) = "" Then PNumTip(i) = "1"
			Next

			sant = "X"
			bseci = 0
			bsecix = 0

			For i=0 To 8
				bsec(i) = 0
				bnum(i) = 0
				bnumG(i) = 0
			Next

			inx = 0

			strsql = " SELECT * FROM CSESTDET "
			strsql = strsql & " where detcon = " & bobj & "  "
			strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
			strsql = strsql & " and not ( "
			strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
			strsql = strsql & " ) and stmodi = '' "

			If session("APLMDCLI")="BC" Then
				strsql=strsql & " AND DETTIPO < 'G'"
			End If
			strsql = strsql & " order by TSEC  "
			' response.write strsql

			Set Rsxx = Server.CreateObject("ADODB.Recordset")
			Rsxx.CursorType = 3
			Rsxx.CacheSize = 100

			Rsxx.Open strsql, cn


			If Not rsxx.eof Then
				ars = rsxx.GetRows()

			Else
				campos = rsxx.fields.count - 1
				ReDim  ars(CInt(campos),0)
				ars(39,0) = "N"
				wentro=1
			End If


			rsxx.Close
			Set rsxx = Nothing

			Do While wentro = 0



				For Inx = inx To UBound(ars,2)



					SCLAVE   = ars(0,inx)
					Stip     = ars(5,inx)

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




						Select Case CStr(spsecn & "")
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
							PFcol  = "#aaaaaa"
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
							Select Case UCase(Pnumcls)
								Case "SIMPLE"
								xPnumS = "SI"
								Case "COMPUESTA"
								xPnumS = "SI"
							End Select

						End If
					End If


					' TITULOS Y SUBTITULOS --------------------------------------------------

					If stip <> "G" And stip <> "H" And stip <> "I"  And stip <> "J" Then
						'Descripcion
						PNumS = ""
						If Pnumer = "SI" Then
							Select Case UCase(Pnumcls)
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



						'''''''''' arma tabla de indices '''''''''''''''''''''''''



						stexto1 = PnumS & Capitalize(ars(7,inx) & "")
						stexto1 = Replace(stexto1," ","&nbsp;")

						For nn=0 To (bseci+1)
							'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
							stexto1 = "&nbsp;" & stexto1
						Next

						wwsNIV2     = bobj & ""
						wwsNIV3     = (ars(0,inx) & "")

						sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

						sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
						iitemC=iitemC+1
						sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
						Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

						CTABELEN = ofv.GenCelda("","utbrr xstyle='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
						FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")


						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




					End If



					' PARRAFOS DE TEXTO    --------------------------------------------------

					If stip = "G" Then

						PNumS = ""
						If Pnumer = "SI" Then
							Select Case UCase(Pnumcls)
								Case "NO"
								BnumG(6) = BnumG(6) - 1

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

							If UCase(Pnumcls) <> "NO" Then

								sTextoV =  ars(8,inx)

								If Len(sTextoV) > 50 Then sTextoV =  Left(sTextoV,50) & "..."

								'''''''''' arma tabla de indices '''''''''''''''''''''''''



								stexto1 = PnumS & Capitalize(sTextoV & "")
								stexto1 = Replace(stexto1," ","&nbsp;")

								For nn=0 To (bseci+1)
									'''    stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
									stexto1 = "&nbsp;" & stexto1
								Next

								wwsNIV2     = bobj & ""
								wwsNIV3     = (ars(0,inx) & "")

								sAccion = xSolaspd & solord(x) & "&NIV1=" & wwsNIV1  & "&NIV2=" & wwsNIV2  & "&NIV3=" & wwsNIV3  & "&NIV4=" & wwsNIV4  & "&NIV5=" & wwsNIV5

								sInput  = "<font onclick=window.top.imenu=this.id; id=ifont" & iitemC  &" style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#bbbbbb;' >" & stexto1 & "</font>"
								iitemC=iitemC+1
								sLink   = OFV.GenLink("", "cmb  xstyle='font-size:8pt;cursor:pointer;text-decoration:none;color:#555555;font-weight:normal;'  ", sAccion, "WK1", "", "", sInput, "si")
								Celda   = Celda & wwwsaltador & "<span ><font color='#ffaa22' size=2 >&#0124;</font>&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

								CTABELEN = ofv.GenCelda("","utbrr xstyle='background-color:#e5e5ff;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#9966dd;padding-left:10px;' ","","","","","",sLink,"si")
								FTABELEN = FTABELEN & ofv.GenRow("","","","",lengtd,"","",CTABELEN,"si")


								'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

				Next

				If inx > UBound(ars,2) Then
					Exit Do

				End If

				ssale = "no"
				Do Until ssale = "si"

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

			rsxx1.movenext
		Loop
		Call ofv.cerrarconsulta(rsxx1)


		Case "12"

		Case "13"

		Case "14"

		Case "15"

		Case "16"



	End Select

	GetIndice = "ok"

End Function

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
			If ci > 0 Then ConvNumer = Mid(CLetra,ci,1)
			ConvNumer = ConvNumer & Mid(CLetra,CR,1)
		End If
		If Ctipo = "A" Then ConvNumer = UCase(ConvNumer)

		Case "I","i"
		CR = CNumero
		Ci = Len(CR)
		If Ci < 1 Or ci > 3 Then
			ConvNumer = "AAAA"
		Else
			cx = 1
			For CL=ci To 1 Step -1
				Select Case cx
					Case 1
					ConvNumer = CRuni(Mid(cr,cl,1))
					Case 2
					ConvNumer = CRdec(Mid(cr,cl,1)) & ConvNumer
					Case 3
					ConvNumer = CRcen(Mid(cr,cl,1)) & ConvNumer
				End Select

				cx = cx + 1
			Next
			If Ctipo = "i" Then ConvNumer = LCase(ConvNumer)
		End If
	End Select


End Function


%>
<script>

var iframeBody = window.top.document;
var iframe = document.getElementById('bodyM');
iframe.addEventListener("mouseover", func, false);
iframe.addEventListener("mouseleave", func2, false);
function func2()
{

if (window.top.CorMoc=='O1'){
 window.top.CorMoc='C';
 document.getElementById('conteM').style.visibility = 'hidden';
  var iframeBody = window.top.document;
   var iframe = iframeBody.getElementById('fxtdoc');
   iframe.cols='50,0,0,*,0,0,0';
 }
}
function func()
{

   var iframeBody = window.top.document;
   var iframe = iframeBody.getElementById('fxtdoc');
   if (window.top.CorMoc=='C'){
     iframe.cols='250,0,0,*,0,0,0';
     window.top.CorMoc='O1';
     document.getElementById('conteM').style.visibility = '';

   }

  // alert(window.top.CorMoc);
}
</script>