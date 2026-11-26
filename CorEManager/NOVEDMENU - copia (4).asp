
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok = ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	shtml = ok
Else
	response.expires = 0
	xxform = ""
'RESPONSE.WRITE Request.ServerVariables("SERVER_NAME")
	SESSION("volver2") = ""
	stab = request.querystring("tab")

	' response.write request.querystring

	If stab = "" Then stab = "view1"
	SESSION("ModMain") = "CL"
	wmnnoved = "NOVED" & wkaplic

	SESSION(wmnnoved) = "SI"
	aplcheck()

	wspviewopc = 2

	stilo1 = " style={background-color:" & SESSION("CMFMENU") & ";border-style:outset;border-width:2;}"
	stilo2 = " style={background-color:" & SESSION("CMFMNBOT") & ";}"


	ReDim Preserve bConsultar(12)

	''''response.write "<br><br><br><br>" & Session("Form") & "<br>"

	Novedform = "CLIMN16317206150930885546227"

	If (SESSION("UN") & "") = "" Then


		xxxx = NEWEMPRE("NO")


		SEMPRE = ATREMP(0, 0)
		SESSION("UN") = SEMPRE & ""

		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1


	End If

    'response.write Request.ServerVariables ("HTTPS")

	Set cn = ofv.conectar(ofv.strconn2)
	Set cn0 = ofv.conectar(ofv.strconn0)

	wmnfuncion = ""

	atmensajes = 0
	Dim atitsecciones(50, 5)
	ttitsecc = 0
	acctitblks = 0
	acctitlin = 0

	Dim atit(50,3)
	Dim awmnasp(50)
	wmntope = 12
	wmntopeX = 12
	WMNINITX = 2

	For tt = 1 To wmntope
		atit(tt,0) = 0
		atit(tt,1) = ""
		atit(tt,2) = "N"
		atit(tt,3) = 0
	Next

	If wkaplic = "OR" Or wkaplic = "PL" Then
		'   wmntopeX = 5
		'   WMNINITX = 4
	End If


	'----------------------------------------------------------------
	'----novedtitex
	titc = 0
	StrSql = "SELECT idsolapa,titulo,activo,cantreg,codEmp FROM novedtitex where  codemp=" & session("UN") & " order by idsolapa "
	'response.write strsql
	Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


	'If Not rsx.EOF Then
	tt=1
	Do While Not rsx.eof


		If rsx("activo")="S" Then

			atit(tt,0) =rsx(0) ' id
			atit(tt,1) =rsx(1) ' titulo
			atit(tt,2) =rsx(2) ' activo  S/N
			atit(tt,3) =rsx(3) ' cant reg a mostrar
			ofv.ObtenerAtributosNOdo Novedform, "", "TIPO", CStr(tt)

			bConsultar(tt) = ofv.AConsultar
			' response.write "tt=" &  tt & "  " & bConsultar(tt) & "<br>"
			If tt = 10 Then

				If bConsultar(tt) Then
					SESSION("SMPLMENU") = "N"
					atit(tt,2) = "N"

				Else
					If SESSION("SMPLMENU") <> "S" Then
						SESSION("SMPLMENU") = "S"
						wmnfuncion = "<script language='javascript' > hdemenu(); </script> "
					End If

				End If

			Else

				If bConsultar(tt) Then
					titc = titc + 1

				Else

					atit(tt,2) = "N"

				End If

			End If

		End If
		titc = titc + 1

		' end if




		rsx.movenext
		tt=tt+1
	Loop

	If atit(10,2) ="S" Then

		ii = 10
		xxxx = wmnbuscamenu("CLIS150172171122331051096140","S")
		wmntope = ii

	End If


	Call ofv.cerrarconsulta(rsx)
	' fin nodedtitex
	'----------------------------------------------------------------


	menuSel = request.querystring("menu")
	Menu = menuSel

	If Menu = "" Then

		If atit(3,2) <> "N" Then
			'menu = 3
			If wkaplic = "OR" Then
				Menu = 4
				menuSel = 4
			Else
				Menu = 3
				menuSel = 3
			End If


		Else
			Menu = 0
			menuSel = 0

			For iix = 1 To 10
				If atit(iix,2) <> "N" Then
					Menu = iix
					menuSel = iix
					Exit For
				End If
			Next
		End If

	End If

	Menu = CInt(Menu)

	'If SESSION("FormN") = "" Then SESSION("FormN") = "NOVEDADES"
	SESSION("FormN") = "NOVEDADES"

	If SESSION("SMPLMENU") = "S" Then SESSION("FormN") = "MENU PRINCIPAL"

	shtml = ofv.FormHeader(SESSION("FormN"))
	shtml = Replace(shtml, "</HEAD><Title>", "<link href=""template6/tabcontent.asp"" rel=""stylesheet"" type=""text/css"" /></HEAD><Title>")

	'' sHTML = ofv.menuheader("","novedades ")

	xs = ""

	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & " function getPosition(obj) {"
	xs = xs & "    var curleft = curtop = 0;"
	xs = xs & "    if (obj.offsetParent) {"
	xs = xs & "        do {"
	xs = xs & "            curleft += obj.offsetLeft;"
	xs = xs & "            curtop += obj.offsetTop;"
	xs = xs & "        } while (obj = obj.offsetParent);"
	xs = xs & "    }"
	xs = xs & "   return curtop;"
   xs = xs & "}"

	xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	If SESSION("APLMDCLI") = "BC" Then
		xs = xs & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If
	xs = xs & " myWin= open(Datos"

	xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom,"    '& Chr(34)
	' xs = xs & "'menubar=yes,toolbar=yes,height=h,width=w,resizable=yes'" '& Chr(34)
	xs = xs & "  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'"    '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

	s = "<script language=JavaScript>" & Chr(13)
	s = s & "function AbrirVentana(Datosx) {" & Chr(13)
	s = s & " var h=570;w=790 ;"
	'If SESSION("APLMDCLI") = "BC" Then
		s = s & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'End If

	s = s & " myWin= open(Datosx"
	s = s & "," & Chr(34) & "edi" & Chr(34) & ","    '& Chr(34)
	'  s = s & "menubar=no,toolbar=no,height=570,width=790,status=yes,resizable=yes" & Chr(34)
	s = s & "  'status=yes,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'"    '& CHR(34)
	s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
	s = s & " } " & Chr(13) & "</script>" & Chr(13)

	shtml = shtml & xs & s & wmnfuncion


     if session("APLMDCLI") <> "BGAL" then
	  shtml = shtml & "<script src=""tabcontent.js"" type=""text/javascript""></script>"
    end if


	tabla1 = ""
	tablab = ""
	tablan = ""
	tablab1 = ""




	nFecini = 0

	If nFecini < CLng(CDate("01-01-1900")) Then
		nFecini = CLng(CDate("01-01-1900"))
	Else
		nFecini = 0
	End If

	dFechai = CLng(Date) - nFecini
	xTrColor = "#FFFFFF"

	sglmar = "<font class=btiW >*</font>"
	sglman = "<font class=btiW >1</font>"
	sglprf = "<font class=btic >&#0128;</font>"
	sgllnk = "<font class=btic >&#0126;</font>"
	sgladj = "<font class=btic >&#0204;</font>"


	Dim ANTabla(500, 25)


	If SESSION("WEBDESK") = "S" Then


		'''''''''''''' buscador por texto     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        IF SESSION("APLMDCLI")="BGAL" THEN
          tablab = tablab & "<DIV id='DTBUSCAR' XSTYLE='float: left; z-index:1;overflow:auto;width:98%;xheight:180px;border-style:solid;border-width:2;border-color:#ffffff;' >"
        ELSE
		 tablab = tablab & "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;xheight:180px;border-style:solid;border-width:2;border-color:#ffffff;' >"
		END IF


		'''  sInputT  = ofv.GenerarInput("textot","","agregar","","80","dt style='border-style:solid;border-width:1;border-color:#666666;' ","")

		If SESSION("APLMDCLI") = "BV" Or SESSION("APLMDCLI") = "BGAL"  Or SESSION("APLMDCLI") = "BNX" Or SESSION("APLMDCLI") = "CORF" Or SESSION("APLMDCLI") = "COR"  Or SESSION("APLMDCLI") = "BC"  Or SESSION("APLMDCLI") = "BP" Or SESSION("APLMDCLI") = "BPN" Or SESSION("APLMDCLI") = "GP" Then
			sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' size='25' style=''width:80%;border-style:solid;border-width:1;border-color:#666666;' >"
		Else
	                sInputT = "<input type='text' class='dt' name='textot' id='textot' size='25' style='width:60%;border-style:solid;border-width:1;border-color:#666666;' >"
		End If

		'''celda = ofv.GenCelda("", "ietx colspan=4", "center valign=middle", "40%", "60", "", "", "<div STYLE=""font-weight:bold;font-family:ARIAL;font-size:24px;color:#ffffff""}>CUERPO NORMATIVO</div>", "si")

        select case  SESSION("APLMDCLI")
        case "BGAL"
                'celda = ofv.GenCelda("", "isttBF ", "left xvalign=middle", "20%", "60", "", "", "Normas&nbsp;&nbsp;&nbsp;", "si")
				celda = ofv.GenCelda("", "isttBF ", "left xvalign=middle", "20%", "60", "", "", "Procesos&nbsp;&nbsp;&nbsp;", "si")

				formx=""'ofv.GenRow("", "", "", "", "", "", "", celda, "si")
				celda = celda &  ofv.GenCelda("", "UT  ", "right", "30%", "", "", "", sInputT, "si")

				sInput = ofv.GenerarInput("", "", "submit", "", "5", "at STYLE='background:url(IMAGENES/LUPA.GIF) no-repeat;color:#ffffff;font-weight:bold;Xbackground-color:#990000;border-style:NONE;border-width:1;Xborder-color:#dddddd;height:30px;WIDTH:30px;cursor:hand;' ", "")
				celda = celda & ofv.GenCelda("", "vt", "", "50%", "", "", "", sInput, "si")

				formx =formx& ofv.GenRow("", "", "", "", "", "", "", celda, "si")
              '    formx=sInputT
				sTarget = ""


				'sAccion  = "consBuscadorSet.asp"
				'sAccion = "SubconsEXTERN.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
				If SESSION("APLMDCLI") = "BV" Or SESSION("APLMDCLI") = "BNX" Or SESSION("APLMDCLI") = "COR" Or SESSION("APLMDCLI") = "BC" Or SESSION("APLMDCLI") = "BGAL" Or SESSION("APLMDCLI") = "BP"   Or SESSION("APLMDCLI") = "BPN" Then
					sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREX&FORMULARIOS=D"
					'sTarget = "_new"
				ElseIf SESSION("APLMDCLI") = "BGALq" Then
					sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOR"
					sTarget = "_new"
				ElseIf SESSION("APLMDCLI") = "CORF" Then
					sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREXF"
				Else
					sAccion = "consBuscadorSet.asp"
				End If

				Fila = ofv.genform("Subconsulta", "", sAccion, sTarget, "post", "", formx, "si")

				tablab = tablab & ofv.GenTabla("", "aaa  xSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", Fila, "si")


        case else
		celda = ofv.GenCelda("", "isttBF colspan=4", "center valign=middle", "40%", "60", "", "", "CUERPO&nbsp;NORMATIVO", "si")


		formx=ofv.GenRow("", "", "", "", "", "", "", celda, "si")
		celda = ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
		formx=formx & ofv.GenRow("", "", "", "", "", "", "", celda, "si")


	         SfrmIndice=""
	         sInputind=""

		' response.write atit(12,2)
		      if atit(12,2)="S" then
	        	' sfuncion="onclick=document.getElementById('frmindice').submit();"

		          sfuncion="onclick=FireEsquemas('FESQ');"

						  	If SESSION("APLMDCLI") = "BPN" Or SESSION("APLMDCLI") = "GP" Then	'Cliente: BPN/GP --> Se modifica el diseño del botón para que sea igual que los demás títulos.
																	'Se crea la clase "btind" para que tome el diseño de la clase

								sInputind =  ofv.GenerarInput("", atit(12,1) & "", "button", "", "5", "btind  colspan='2' " & sfuncion, "")

								Else

									sInputind =  ofv.GenerarInput("", atit(12,1) & "", "button", "", "5", "bietx " & sfuncion, "")

							End If

              end if



						shref="subconsulta.asp?MENU=PORTEMA&NOMBRE=TEMA&NDAT=ESTADO&NVAL=2&seg=CLIMS154172171124031051096140&MNN=POR_INDICE&base=D"
						sInput = ofv.GenerarInput("", "Buscar", "submit", "", "5", "bietx xSTYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")


				If SESSION("APLMDCLI") = "BPN" Or SESSION("APLMDCLI") = "GP" Then	'Cliente: BPN/GP --> Se modifica el diseño para que primero figure el cuadro de texto
														'y el botón del buscador y abajo el botón "Indice"

					celda = celda & ofv.GenCelda("", "ut", "center", "90%", "", "", "", sInputT & "&nbsp;&nbsp;"& sInput, "si")
					formx =formx& ofv.GenRow("", "", "", "", "", "", "", celda, "si")

					celda =  ofv.GenCelda("", "UT colspan='2' ", "center", "30%", "", "", "", "<br>" & sInputind & "&nbsp;", "si")
					formx =formx& ofv.GenRow("", "", "", "", "", "", "", celda, "si")

					Else

						celda =  ofv.GenCelda("", "UT' ", "right", "30%", "", "", "", sInputind & "&nbsp;", "si")
						'formx =formx& ofv.GenRow("", "", "", "", "", "", "", celda, "si")

						celda = celda & ofv.GenCelda("", "ut", "", "70%", "", "", "", sInputT & "&nbsp;&nbsp;"& sInput, "si")
						formx =formx& ofv.GenRow("", "", "", "", "", "", "", celda, "si")

				End If

		sTarget = ""


		'sAccion  = "consBuscadorSet.asp"
		'sAccion = "SubconsEXTERN.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		If SESSION("APLMDCLI") = "BV" Or SESSION("APLMDCLI") = "BNX" Or SESSION("APLMDCLI") = "COR" Or SESSION("APLMDCLI") = "BC" Or SESSION("APLMDCLI") = "BGAL" Or SESSION("APLMDCLI") = "BP"   Or SESSION("APLMDCLI") = "BPN" Then
			sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREX&FORMULARIOS=D"
			'sTarget = "_new"
		ElseIf SESSION("APLMDCLI") = "BGALq" Then
			sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOR"
			sTarget = "_new"
		ElseIf SESSION("APLMDCLI") = "CORF" Then
			sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREXF"
		Else
			sAccion = "consBuscadorSet.asp"
		End If

		Fila = ofv.genform("Subconsulta", "", sAccion, sTarget, "post", "", formx, "si")


		tablab = tablab & ofv.GenTabla("", "aaa xSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "0", "0", "", "", Fila, "si")
        tablab=tablab & SfrmIndice
        end select
		tablab = tablab & "</DIV>" & Chr(13)

		tablabcc=tablab

		'ALTA NORMATIVA

		'If SESSION("APLMDCLI") = "BGAL" Then

		If session("CORMODUSLTNORM")= "S" Then

			tablab = tablab & "<DIV>"


			'sAccion = "crearNormativa.asp?pedido=ALTA"
			aspeditar = "Editar.asp?tipocpt=SN&ACC=SSN&DOCMODAL=CL&pedido=ALTA"
			'aspeditar = "crearNormativa.asp?pedido=ALTA"
			'aspeditar = CStr(aspeditar)
			aspeditar = "EDITARDOC.asp?acc=SNN&pedido=ALTA&seg=" & seg &"&docmodal=CL"
			sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

			IF session("APLMDCLI")="BGAL" THEN
			    'agilita com light
					LINKBSQ2 = "<input type=button style='background-color:#f4f4f4;width:300px;height:40px;font-family:calibri;font-size:12pt;color:#000000;'" & sFuncion & " value='Generar Solicitud ABM Procesos' >"
					ELSE
					LINKBSQ2=" <input type=button  class=bietx xstyle='cursor:hand;'" & sFuncion & " value='Generar Solicitud Normativa ' >"
			end if


			' LINKBSQ2=" <input type=button  xclass=bietx xstyle='cursor:hand;' xSTYLE='width:90%;height:50px;color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' " '& sFuncion & " value='Generar Solicitud ABM Norma' >"
			' LINKBSQ2 = LINKBSQ2 & "style='background-color:#f4f4f4;width:300px;height:40px;font-family:calibri;font-size:12pt;color:#000000;'" & sFuncion & " value='Generar Solicitud ABM Norma' >"
            LINKBSQ2BG=LINKBSQ2
			ofv.ObtenerAtributos "CLIMN54011722041514609825499",""

			bAconsultar = ofv.Aconsultar

			If Not bAconsultar Then
				LINKBSQ2=""
				LINKBSQ2BG=""
			End If

			'FormNorm = ofv.genform("FormNorm","",sAccion,"_new","post","",sInput, "si")
			celda = ofv.GenCelda("", "vt", "", "", "", "", "", LINKBSQ2, "si")
			fila = ofv.GenRow("", "", "", "", "", "", "", celda, "si")
			tablab = tablab & ofv.GenTabla("", "a ", "center", "", "", "0", "", "-2", "0", "", "", Fila, "si")

			tablab = tablab & "</DIV>"
			''''''''''''''''''''''''

		End If

        IF session("APLMDCLI")="BGAL" THEN
         tablabFORM=  "<DIV id='DTBUSCAR2' XSTYLE='float: left; z-index:1;overflow:auto;width:98%;height:100%;border-style:solid;border-width:2;border-color:#ffffff;' >"
        ELSE
		 tablabFORM=  "<DIV id='DTBUSCAR2' STYLE='float: left; z-index:1;overflow:auto;width:98%;height:100%;border-style:solid;border-width:2;border-color:#ffffff;' >"
		END IF
		celda=""
		fila=""

		'''  sInputT  = ofv.GenerarInput("textot","","agregar","","80","dt style='border-style:solid;border-width:1;border-color:#666666;' ","")

		If  session("APLMDCLI")="COR"  Or session("APLMDCLI")="BGAL" Or session("APLMDCLI")="BC"   Then
			sInputT = "<input type='text' class='dt' name='buscar' id='textot' size='30' style='border-style:solid;border-width:1;border-color:#666666;'  >"
			sInputT2 = "<input type='text' class='dt' name='buscar' id='textot' size='30' style='border-style:solid;border-width:1;border-color:#666666;' onkeyup=mascara(this,'-',patron,true) >"
		Else
			sInputT = "<input type='text' class='dt' name='buscar' id='textot' size='30' style='border-style:solid;border-width:1;border-color:#666666;' >"
		End If

		existeF=wmnbuscamenu("CLIS245159671557361072846976","N")  ' busco que tenga menu formularios
		'response.write session("CORMODUSOLABM") & "  "   &  existeF

		If  existeF="OK" Then
			If session("CORMODUSOLABM") <> "S"  Then existeF="NO"
		End If

		' if session("CORMODUSOLABM") <>"S"  then existeF="NO"
		celda=""
		If existeF="OK"  Then

		 select case  session("APLMDCLI")

		 Case "BGAL"
		    celda = ofv.GenCelda("", "isttBF ", "left vvalign=middle", "20%", "60", "", "", "Formularios", "si")
			''Fila = ofv.GenRow("", "", "", "", "", "", "", celda, "si")
			'celda=ofv.GenCelda("", "xxx colspan=2", "", "50% height=10px", "", "", "&nbsp;", "", "si")
			'Fila = fila & ofv.GenRow("", "", "", "", "", "", "", celda, "si")
			'celda=ofv.GenCelda("", "UT  ", "LEFT", "1%", "", "", "", "" , "si")
			celda =  CELDA & ofv.GenCelda("", "UT  ", "RIGHT", "50%", "", "", "",   sInputT, "si")
			' celdaX=ofv.GenCelda("", "UT  ", "LEFT", "1%", "", "", "", "CODIGO " , "si")
			' celdaX = CELDAX &  ofv.GenCelda("", "UT  ", "LEFT", "30%", "", "", "",  sInputT2, "si")

			'sInput = ofv.GenerarInput("", ".", "submit", "", "5", "at STYLE='background'':url(/imagenes/lupa.gif) ;color:#ffffff;font-weight:bold;Xbackground-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")
			SINPUT="<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('Subconsultaform').submit();"">"
			'SINPUT="<font  face=webdings  size=3 color=#ffffff onclick=""document.getElementById('Subconsultaform').submit();"">L</font>"
			SINPUT2="" '"<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('Subconsultaform2').submit();"">"
			celda = celda & ofv.GenCelda("", "vt", "", "30%", "", "", "", sInput, "si")
			SHREF="SCONSUL.ASP?Ing=P&BASE=K&seg=CLIMN51812732031257124269088&MNN=CONSULTA_GENERAL_DE_FORMULARIOS"
			' LINKBSQ = ofv.GenLink("", "istt ", sHref, "", "", "", "BUSQUEDA<br>AVANZADA", "si")
			LINKBSQ1=" <input type=button  class=bietx style='cursor:hand;'	title='Busqueda Avanzada' xSTYLE='width:90%;height:50px;color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' onclick=""document.getElementById('Subconsultaformex').submit();"" value='Busqueda Avanzada'>"
			sTarget = "info"
			'//////////// agregado a pedido de Galicia, anula busqueda avanzada para todos  ///////////
			LINKBSQ1="&nbsp;"
			'//////////// agregado a pedido de Galicia, anula busqueda avanzada para todos  ///////////
			xfrm=ofv.GenForm("Subconsultaformex", "", SHREF, sTarget, "post", "", "", "si")

			'celda = celda & ofv.GenCelda("", "zzzz rowspan=2 align=left", "", "50%", "", "", "", "" , "si")
			celdaX =  ofv.GenCelda("", "vt", "", "5%", "", "", "", sInput2, "si")
			formx = ofv.GenRow("", "", "", "", "", "", "", celda, "si")
			'formx="<table border>" & formx & "</table>"
			'formx2 = ofv.GenRow("", "", "", "", "", "", "", celdaX, "si")

		 Case else
			celda = ofv.GenCelda("", "isttBF colspan=4", "center valign=middle", "40%", "60", "", "", "BASE DE FORMULARIOS", "si")
			Fila = ofv.GenRow("", "", "", "", "", "", "", celda, "si")
			celda=ofv.GenCelda("", "xxx colspan=2", "", "50% height=10px", "", "", "&nbsp;", "", "si")
			Fila = fila & ofv.GenRow("", "", "", "", "", "", "", celda, "si")
			celda=ofv.GenCelda("", "UT  ", "LEFT", "1%", "", "", "", "" , "si")
			celda =  CELDA & ofv.GenCelda("", "UT  ", "LEFT", "30%", "", "", "",   sInputT, "si")
			' celdaX=ofv.GenCelda("", "UT  ", "LEFT", "1%", "", "", "", "CODIGO " , "si")
			' celdaX = CELDAX &  ofv.GenCelda("", "UT  ", "LEFT", "30%", "", "", "",  sInputT2, "si")

			'sInput = ofv.GenerarInput("", ".", "submit", "", "5", "at STYLE='background'':url(/imagenes/lupa.gif) ;color:#ffffff;font-weight:bold;Xbackground-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")
			SINPUT="<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('Subconsultaform').submit();"">"
			'SINPUT="<font  face=webdings  size=3 color=#ffffff onclick=""document.getElementById('Subconsultaform').submit();"">L</font>"
			SINPUT2="" '"<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('Subconsultaform2').submit();"">"
			celda = celda & ofv.GenCelda("", "vt", "", "5%", "", "", "", sInput, "si")
			SHREF="SCONSUL.ASP?Ing=P&BASE=K&seg=CLIMN51812732031257124269088&MNN=CONSULTA_GENERAL_DE_FORMULARIOS"
			' LINKBSQ = ofv.GenLink("", "istt ", sHref, "", "", "", "BUSQUEDA<br>AVANZADA", "si")
			LINKBSQ1=" <input type=button  class=bietx style='cursor:hand;'	title='Busqueda Avanzada' xSTYLE='width:90%;height:50px;color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' onclick=""document.getElementById('Subconsultaformex').submit();"" value='Busqueda Avanzada'>"
			sTarget = "info"
			'//////////// agregado a pedido de Galicia, anula busqueda avanzada para todos  ///////////
			LINKBSQ1="&nbsp;"
			'//////////// agregado a pedido de Galicia, anula busqueda avanzada para todos  ///////////
			xfrm=ofv.GenForm("Subconsultaformex", "", SHREF, sTarget, "post", "", "", "si")

			celda = celda & ofv.GenCelda("", "zzzz rowspan=2 align=left", "", "50%", "", "", "", "" , "si")
			celdaX = celdaX & ofv.GenCelda("", "vt", "", "5%", "", "", "", sInput2, "si")
			formx = ofv.GenRow("", "", "", "", "", "", "", celda, "si")
			formx2 = ofv.GenRow("", "", "", "", "", "", "", celdaX, "si")
		End Select


		Else
			formx=""
			formx2=""

		End If
		'

		OFV.ObtenerAtributosNodo "CLIMN51812732031257124269088", "", "ESTADO", "3"
		bfconsultar = OFV.AConsultar

		'  response.write  "<br>" & bfconsultar & "  "  & existeF

		If   Not bfconsultar Or  existeF="NO" Then
			LINKBSQ1=""

		End If



		'     ofv.ObtenerAtributos "CLIMN5261261104010256728451",""


		SHREF="PedidoForm.ASP?Ing=P&BASE=K&seg=CLIMN52412655104952828166469"


		' sHref = "&nbsp;"
		'
		sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & SHREF & "'); return false;" & Chr(34)

		'LINKBSQ = ofv.GenLink("", "istt ", sHref, "", "", sFuncion, "<B>GENERAR PEDIDO</B>", "si")
		' LINKBSQ = ofv.GenLink("", "ZZZ", sHref, "", "", "", "<B>GENERAR PEDIDO</B>" , "si")
		If session("CORMODUSOLABM")="S" Then

			IF session("APLMDCLI")="BGAL" THEN
				LINKBSQ2=" <input type=button  style='background-color:#f4f4f4;width:300px;height:40px;font-family:calibri;font-size:12pt;color:#000000;'" & sFuncion & " value='Generar Solicitud ABM Formulario' >"
				ELSE
					LINKBSQ2=" <input type=button  class=bietx xstyle='cursor:hand;'	title='Generar Solicitud ABM Formulario' " & sFuncion & " value='Generar Solicitud ABM' >"
			end if

			' LINKBSQ2=" <input type=button  class=xbietx xstyle='cursor:hand;'	title='Generar Solicitud ABM Formulario' xSTYLE='width:90%;height:20px;color:#ffffff;xfont-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:50px;cursor:hand;' " '& sFuncion & " value='Generar Solicitud ABM Formulario' >"
			' LINKBSQ2 = LINKBSQ2 & "style='background-color:#f4f4f4;width:300px;height:40px;font-family:calibri;font-size:12pt;color:#000000;'" & sFuncion & " value='Generar Solicitud ABM Formulario' >"
			' 'background-color:#990000'
		Else
			LINKBSQ2=""
		End If
		ofv.ObtenerAtributos "CLIMN5261261104010256728451",""
		' response.write ofv.Aconsultar
		bAconsultar = ofv.Aconsultar
		If   Not bAconsultar  Then
			LINKBSQ2=""
		End If


		'xfrm=ofv.GenForm("Subconsultaformex", "", SHREF, sTarget, "post", "", "", "si")
        if session("APLMDCLI") <> "BGAL" THEN
		 celdaX3 = ""' ofv.GenCelda("", "vti COLspan=2 align=center  HEIGHT=40PX", "", "100%", "", "", "", LINKBSQ1, "si")
		' celdaX3 = celdaX3 & ofv.GenCelda("", "ut", "", "100%", "", "", "", "&nbsp;", "si")
		 celdaX3 =  celdaX3 & ofv.GenCelda("", "ut COLspan=3 align=center HEIGHT=40PX", "", "100%", "", "", "", LINKBSQ2, "si")
		 formx3 = ofv.GenRow("", "", "", "", "", "", "", ofv.GenCelda("", "zzz COLspan=5", "", "100%", "", "", "", "&nbsp;", "si"), "si")
		 formx4 = ofv.GenRow("", "", "", "", "", "", "", celdaX3, "si")
		 formx5 = ofv.GenRow("", "", "", "", "", "", "", ofv.GenCelda("", "zzz COLspan=5", "", "100%", "", "", "", "&nbsp;", "si"), "si")
       END IF

		sTarget = ""


		If   bfconsultar And existef="OK" Then

			'sAccion  = "consBuscadorSet.asp"
			'sAccion = "SubconsEXTERN.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
			If session("APLMDCLI") = "COR" Or session("APLMDCLI")="BC"   Or session("APLMDCLI")="BGAL" Then
				sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREX&FORMULARIOS=K"
				'sTarget = "_new"
			ElseIf session("APLMDCLI") = "BGALq" Then
				sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOR"
				sTarget = "_new"
			Else
				sAccion = "consBuscadorSet.asp"
			End If

			Fila = Fila & ofv.GenForm("Subconsultaform", "", sAccion, sTarget, "post", "", formx, "si")

			'	If  session("APLMDCLI")="BGAL" Then
			sAccion="Conscodigo.asp"
			sTarget = "fesquemas"
			Fila = Fila & ofv.GenForm("Subconsultaform2", "", sAccion, sTarget, "post", "", formx2, "si")
			Fila = Fila & formx3 & formx4 & formx5
			'	End If
            if session("APLMDCLI")="BGAL" THEN
			 tablabFORM = tablabFORM & "" & ofv.GenTabla("", "aaa     xSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", Fila, "si")
			ELSE
			 tablabFORM = tablabFORM & "<BR><BR>" & ofv.GenTabla("", "aaa     xSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", Fila, "si")
			END IF

		End If
		tablabFORM = tablabFORM & "</DIV>" & xfrm & Chr(13)
		'tablab =tablab &" <script language='javascript'>document.Subconsulta.submit();</script>"


		tablaLN=""
		If  session("APLMDCLI")="BPN"  Then
			''''''''''''''''''''''''''CONSULTA LINEA DE NEGOCIO

			tablaLN = tablaLN & "<div><br>"

			tablaLN = tablaLN & "<table style='height:100%;width:100%;xborder:solid 1px grey;' align=center>"
			tablaLN = tablaLN & "<tr>"
			tablaLN = tablaLN & "<td></td>"
			tablaLN = tablaLN & "</tr>"
			tablaLN = tablaLN & "<tr>"
			tablaLN = tablaLN & "<td></td>"
			tablaLN = tablaLN & "</tr>"
			tablaLN = tablaLN & "<tr>"
			tablaLN = tablaLN & "<td class='ietx' sstyle='color:#000000;'>CONSULTA LINEA DE NEGOCIO</td>"

			tablaLN = tablaLN & "</tr>"
			tablaLN = tablaLN & "<tr>"
			tablaLN = tablaLN & "<td></td>"
			tablaLN = tablaLN & "</tr>"

			sql = "select distinct a.objid,a.descri,a.secuen from tipoproducto a"
			sql= sql & " inner join objetos b on (a.objid = b.tsegmento)  "
			sql= sql & " inner join objprty c on (b.codobjeto = c.codobjeto and b.tipobjeto=c.tipobjeto and c.estado=3)"
			sql= sql & " order by a.secuen"
			'sql= sql & "'response.write sql
			Set rx = ofv.crearconsultaEx(sql,cn,1,parametros)
			xTrColor = "#eeeeff"
			Do While Not rx.eof

				parametros = "MENU=PORLINEGO&NOMBRE=LINEA&NDAT=ESTADO&NVAL=2&seg=CLIMN4601201011384966908594&MNN=POR_LINEA_DE_NEGOCIO&id=" & rx(0)

				tablaLN = tablaLN & "<tr bgcolor=" & xTrColor & " height=25>"


				tablaLN = tablaLN & "<td><a CLASS=nt  href='SUBCONSULTA.asp?" & parametros & "'> -" & rx(1) & "</a></td>"
				tablaLN = tablaLN & "</tr>"

				rx.movenext

				If xTrColor = "#FFFFFF" Then
					xTrColor = "#eeeeff"
				Else
					xTrColor = "#FFFFFF"
				End If

			Loop

			tablaLN = tablaLN & "</table>"

			tablaLN = tablaLN & "</div>"
		End If

		'''''''''''''' indice buscador normativo '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


		tabla1 = tabla1 & "<div id='DTTEMA' style='float: left;width:98%; " & Chr(13)
		tabla1 = tabla1 & "border-style:none;border-width:1;border-color:#000000;display:none;' >"





		tabla1 = tabla1 & "<table width=100% cellspacing=0 cellpadding=1  style='border-style:none;border-width:1;border-color:#666666;' >"
		tabla1 = tabla1 & "<tr xHEIGHT=25  >"

			' If session("APLMDCLI")  <> "BPN" THen		'Cliente <> BPN/GP --> Se saca el título "Indice Normativo" que está fijo.
														' 'El títúlo del botón es el título del índice
														' 'Lo toma de Configuración/Títulos de Novedades

				' xsclase = " CLASS=ietx "
				' '''tabla1 = tabla1 & "<td " & xsclase & "  >" & Capitalize("INDICE NORMATIVO") & "</TD>"
				' tabla1 = tabla1 & "<td " & xsclase & "  >" & "INDICE NORMATIVO " &  "</TD>"
				' tabla1 = tabla1 & "</TR>"
				' tabla1 = tabla1 & "<tr HEIGHT=25  >"




				' If session("APLMDCLI")  <> "GP" tHEN
										' xsclase = " CLASS=ietx "
										' '''tabla1 = tabla1 & "<td " & xsclase & "  >" & Capitalize("INDICE NORMATIVO") & "</TD>"
										' tabla1 = tabla1 & "<td " & xsclase & "  >" & "INDICE NORMATIVO " &  "</TD>"
										' tabla1 = tabla1 & "</TR>"
										' tabla1 = tabla1 & "<tr HEIGHT=25  >"
				' End if

			' End IF

		If (SESSION("APLMDCLI") = "BPN" OR SESSION("APLMDCLI") = "GP") THEN
			tabla1 = tabla1 & "<td class=dt style=font-size:10pt; align=center >-------------------&nbsp;TEMAS&nbsp;------------------- &nbsp;&nbsp;&nbsp;&nbsp;<button  onclick=ClosEsquemas('FESQ');  style='color:red;border-style:solid;'  >X</button></TD>"
			tabla1 = tabla1 & "</TR>"
			tabla1 = tabla1 & "</table>" & Chr(13)

			ELSE

				xsclase = " CLASS=ietx "
				tabla1 = tabla1 & "<td " & xsclase & "  >" & "INDICE NORMATIVO " &  "</TD>"
				tabla1 = tabla1 & "</TR>"
				tabla1 = tabla1 & "<tr HEIGHT=25  >"
				tabla1 = tabla1 & "<td class=dt style=font-size:10pt; align=center >-------------------&nbsp;TEMAS&nbsp;------------------- &nbsp;&nbsp;&nbsp;&nbsp;<button  onclick=ClosEsquemas('FESQ');  style='color:red;border-style:solid;'  >X</button></TD>"
				tabla1 = tabla1 & "</TR>"
				tabla1 = tabla1 & "</table>" & Chr(13)

		END IF



		''''height:400px;

		tabla1 = tabla1 & "<div id='DTTEMAINT' style='overflow:auto;" & wstyl & ";"
		tabla1 = tabla1 & "border-style:solid;border-width:1;border-color:#dddddd;' >" & Chr(13)

		WSPFRMN = "DTTEMAITF"
                WSPDIVN = "DTTEMAINT"
                WSPDIVNM = "DTTEMA"
		'WSPFRMA = "consdocumindex.asp"
		WSPFRMA = "CONSTEMAIND.asp?FORMATO=ABR"

		WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='xdisplay:none;' METHOD='POST' ><input type=button onclick=FireEsquemas('FESQ'); style='display:none;' ></FORM>"

		WSPDTTSCP = ""
		       ''''''WSPDTTSCP = WSPDTTSCP & "<script FOR=window EVENT=onload  >"
		    WSPDTTSCP = WSPDTTSCP & "<script language=javascript  >"
		     WSPDTTSCP = WSPDTTSCP & " function FireEsquemas(FESQ) { "
		     ' WSPDTTSCP = WSPDTTSCP &  " alert('ENTRO: ');"
		     '  WSPDTTSCP = WSPDTTSCP &  "alert(document.getElementById(" & chr(34) & WSPFRMN & chr(34) & ").id);"

		     WSPDTTSCP = WSPDTTSCP &  "document.getElementById(" & chr(34) & WSPDIVNM & chr(34) & ").style.display="  & chr(34) & "block" & chr(34) & ";"
		     ' WSPDTTSCP = WSPDTTSCP &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").style.display);"

		     WSPDTTSCP = WSPDTTSCP & "document." & WSPFRMN & ".submit();"
		     WSPDTTSCP = WSPDTTSCP & " } "
		     WSPDTTSCP = WSPDTTSCP & " function ClosEsquemas(FESQ) { "
		     WSPDTTSCP = WSPDTTSCP & "    document.getElementById(" & chr(34) & WSPDIVNM & chr(34) & ").style.display="  & chr(34) & "none" & chr(34) & "; "
		     WSPDTTSCP = WSPDTTSCP & " } "
		    WSPDTTSCP = WSPDTTSCP & "</script>"

		'oculto indice normativo
		'tabla1=""
		'WSPDTTSCP=""
		'TABLA1 = Tabla1 & "<script src='../coremanager/consdocumindex.asp?sid=1&target=_self' ></script>"


		tabla1 = tabla1 & "</DIV>"
		tabla1 = tabla1 & "</DIV>" & WSPFORME & Chr(13)


		tabla1f="" 'armapedidos()

		If bfconsultar And existef="OK" Then

			tabla1f =  tabla1f &  "<div id='DTTEMA' style='float: left;width:98%; " & Chr(13)
			tabla1f = tabla1f & "border-style:solid;border-width:2;border-color:#ffffff;' >"

			tabla1f = tabla1f & "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
			tabla1f = tabla1f & "<tr HEIGHT=25  >"
			xsclase = " CLASS=ietx "
			''''tabla1 = tabla1 & "<td " & xsclase & "  >" & Capitalize("INDICE NORMATIVO") & "</TD>"
			tabla1f = tabla1f & "<td " & xsclase & "   >" & atit(11,1)& "</TD>"
			tabla1f = tabla1f & "</TR>"

			tabla1f = tabla1f & "</table>" & Chr(13)

			''''height:400px;

			tabla1f = tabla1f & "<div id='DTTEMAINTF' style='overflow:auto;" & wstyl & ";"
			tabla1f = tabla1f & "width:100%;border-style:solid;border-width:1;border-color:#dddddd;' >" & Chr(13)

			tabla1f = tabla1f & ULTFORMCONS("N")




			tabla1f = tabla1f & "</DIV>"
			tabla1f = tabla1f & "</DIV>"  & Chr(13)
		Else
			tabla1f=""
		End If


		XTablaN = ""




		xTrColor = "#FFFFFF"

		XCANT = 0


		'response.write time

		For ii = 2 To wmntope
			'response.write  atit(ii,1)  & "  "  & ii & " <br>"
			If atit(ii,2) <> "N" Then
				XCANT = XCANT + 1

				TNTabla = 0
				INTabla = 0
				xnohay1 = "S"
				xnohay2 = "S"
				saleSel = "N"

				If ii = 10 Then
					stilo = stilo2
				Else
                   if  session("APLMDCLI")  = "BGAL"  then
                    if ii=4 then  ' lo hace abajo autorizaciones
                    else
					   xxxx = noveselec(ii)
				    end if
				   else
				     xxxx = noveselec(ii)
				   end if
					'response.write  session("APLMDCLI")

					TNTabla = INTabla

					If II=4  and  session("APLMDCLI")  <> "BGAL"  Then
						STRAUTORIZACION=noveformatEX(ii, Menu,"N","")
					ElseIf 	II=5  and  session("APLMDCLI")  <> "BGAL" Then
						ultimosDocConsultados=noveformatEX(ii, Menu,"S","")
					ElseIf II=7 and  session("APLMDCLI")  <> "BGAL" Then
						STRDOCVENCER=noveformatEX(ii, Menu,"N","")
					ElseIf II=11 Then
						'STRDOCVENCER=noveformatEX(ii, Menu,"N","")
					Else
						XSOPTABN = noveformat(ii, Menu)
					End If

					'  TABLAN = TablaN &  XSOPTABN

				End If

			End If
		Next


		If ttitsecc > 0 Then


			XTablaN = ""

			XTABLAN2 = ""
			' xsclase = " CLASS=VTI STYLE='background-color:#6666ff;color:#ffffff;font-size:9pt;' "

			xsclase = " class=iett "
			XTablaN = XTablaN & "<div id='DTNOVSEC' style='float: left; "
			XTablaN = XTablaN & "border-style:solid;border-width:2;border-color:#ffffff;' >"


			XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
			XTablaN = XTablaN & "<tr HEIGHT=25  >"
			'''''''XTablaN = XTablaN & "<td " & xsclase & "   >" & Capitalize("secciones") & "</TD>"
			XTablaN = XTablaN & "<td " & xsclase & "   >" & "SECCIONES" & "</TD>"
			XTablaN = XTablaN & "</TR>"
			XTablaN = XTablaN & "</table>"

			XTablaN = XTablaN & "<div id='DTNOVINTSEC' style='"
			XTablaN = XTablaN & "border-style:solid;border-width:1;border-color:#dddddd;' >"

			XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=0  >"

			xxfondo = ""

			colsumd = 0
			colsumi = 0
			ctotspand = 0
			ctotspani = 0
			ctotspanall = 0

			For ii = 1 To ttitsecc



				cworka = CInt(ii / 2)
				cworka = cworka * 2

				If cworka = ii Then
					colsumd = colsumd + atitsecciones(ii, 2)
				Else
					colsumi = colsumi + atitsecciones(ii, 2)
				End If



				XTablaN = XTablaN & "<tr HEIGHT='25'  " & xxfondo & " >"
				XTablaN = XTablaN & "<td class=ut width=1% >&nbsp;<B><font color='#f7aa22' size=2 >&#0149;</font></B></TD>"
				XTablaN = XTablaN & "<td class=ut STYLE='color:#000000;font-size:8pt;font-weight:500;' >&nbsp;"
				XTablaN = XTablaN & ""
				XTablaN = XTablaN & atitsecciones(ii, 0) & "&nbsp;(" & atitsecciones(ii, 1) & ")" & ""
				XTablaN = XTablaN & "</TD>"
				XTablaN = XTablaN & "</TR>"

				If xxfondo = "" Then
					xxfondo = " bgcolor='#ddccff' "
				Else
					xxfondo = ""
				End If



			Next







			XTablaN = XTablaN & "<tr HEIGHT='20' >"
			XTablaN = XTablaN & "<td class=UT colspan=2  >&nbsp;</TD>"
			XTablaN = XTablaN & "</TR>"

			If atmensajes > 99990 Then

				XTablaN = XTablaN & "<tr HEIGHT='20' bgcolor='#f0eedd' >"
				XTablaN = XTablaN & "<td class=vTi colspan=2 STYLE='color:#000000;font-size:8pt;"
				XTablaN = XTablaN & "border-style:solid;border-width:1;border-color:#666666;' >"
				'''''XTablaN = XTablaN & Capitalize("mensajes") & "</TD>"
				XTablaN = XTablaN & "MENSAJES" & "</TD>"
				XTablaN = XTablaN & "</TR>"

				XTablaN = XTablaN & "<tr HEIGHT='25' >"
				XTablaN = XTablaN & "<td class=ut width=1% >&nbsp;<B><font color='#990000' size=3 >!</font></B></TD>"
				XTablaN = XTablaN & "<td class=ut STYLE='color:#000000;' >&nbsp;"
				XTablaN = XTablaN & "<font color=#666666 style='font-size:8pt;font-weight:700;' >"
				XTablaN = XTablaN & "Ud. tiene&nbsp;" & atmensajes & "&nbsp;mensajes&nbsp;pendientes." & "</font>"
				XTablaN = XTablaN & "</TD>"
				XTablaN = XTablaN & "</TR>"

			End If


			XTablaN = XTablaN & "</table>"
			XTablaN = XTablaN & "</DIV>"
			XTablaN = XTablaN & "</DIV>"

			tablab1 = XTablaN

		End If





	Else

		TNTabla = 0
		INTabla = 0
		xnohay1 = "S"
		xnohay2 = "S"
		saleSel = "N"

      ' response.write menu & "wwwwwwwwwwwwwwwww<br>"
		xxxx = noveselec(Menu)

		TNTabla = INTabla


		tablan = ""
		tablan = tablan & "<DIV id='DTNOVE' STYLE='display:BLOCK;z-index:1;margin-top:1;margin-left:1;overflow:auto;'>"


		tablan = tablan & "<table  width=100% cellspacing=0 cellpadding=2 class='HX1CL'  style='border-style:solid;border-width:1;border-color:#ffffff;"
		tablan = tablan & "background-color:#cccccc;"
		tablan = tablan & "' >"

		xnewmenu = ""
		xoldmenu = ""

		tablan = tablan & "<tr >"
		tablan = tablan & "<td class='UT'  >"

		tablan = tablan & "<table   cellspacing=1 cellpadding=3  >"

		For ii = 2 To wmntope
			If atit(ii,2) <> "N" Then

				If ii = 10 Then
					stilo = stilo2
				Else

					If ii = Menu Then
						xsclase = "ttr STYLE='border-style:inset;border-color:" & SESSION("CMFPNTITFT") & ";border-width:1;padding-left:5px;padding-right:5px;" & XspEx("#f7aa22", "#dddddd", 0) & ";' "
					Else
						xsclase = "ttr STYLE='border-right-style:solid;border-right-color:" & SESSION("CMFPNTITFT") & ";border-right-width:1;padding-left:5px;padding-right:5px;' "
					End If

					xmDato = Replace(atit(ii,1), " ", "&nbsp;")

					If ii < 11 Then
						xmhref = "<a class='xts' href='NOVEDMENU.asp?menu=" & ii & "' >" & xmDato & "</a>"
						xoldmenu = xoldmenu & "<td class=" & xsclase & " align=center  >"
						xoldmenu = xoldmenu & xmhref & "</TD>"
					Else
						xmhref = "<a class='xts' href='" & awmnasp(ii) & "' >" & xmDato & "</a>"
						xnewmenu = xnewmenu & "<td class=" & xsclase & " align=center  >"
						xnewmenu = xnewmenu & xmhref & "</TD>"

					End If



				End If

			End If
		Next

		tablan = tablan & xnewmenu & xoldmenu
		tablan = tablan & "</table>"

		tablan = tablan & "</TD>"
		tablan = tablan & "</TR>"
		tablan = tablan & "</table>"

		'''''''''''''''''''''''''''''
		tablan = tablan & "<center>"

		tablan = tablan & "<div id='WMN0' style='overflow:auto;height:50%;" & wstyl & "' >"

		If xnohay1 = "S" And xnohay2 = "S" Then

			tablan = tablan & "<table  margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"
			tablan = tablan & "<tr  height=200 >"
			tablan = tablan & "<td class=ttt width=100% align=center valign=middle ><big>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(Menu,1) & "&nbsp;</TD>"
			tablan = tablan & "</TR>"

		Else
			tablan = tablan & "<table bgcolor=#eeeeee margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"

			For INTabla = 1 To TNTabla

				xtmcod = ANTabla(INTabla, 0)
				xtipom = ANTabla(INTabla, 1)
				xtpri = ANTabla(INTabla, 2)
				xtfvig = ANTabla(INTabla, 3)
				xthvig = ANTabla(INTabla, 4)
				xtsadj = ANTabla(INTabla, 5)
				xtslnk = ANTabla(INTabla, 6)
				xfromusr = ANTabla(INTabla, 7)
				xtmsg = ANTabla(INTabla, 8)
				xtmsgimg = ANTabla(INTabla, 9)
				xtmsglnk = ANTabla(INTabla, 10)

				xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsg)



				Select Case xtipom
					Case 1
					xtipomes = "GENERAL"
					Case 2
					xtipomes = "PARTICULAR"
					Case 3
					xtipomes = "PUBLIC"
					Case 4
					xtipomes = "AUTORIZ"
					Case 5
					xtipomes = "WORKFLOW"


				End Select

				xfclas = "dt"
				sImage = "&nbsp;"

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

				sglmod = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
				sImage = ofv.GenLink("", xfclas & " title='" & sgltit & "' ", "", "", "", "", sglmod, "si")

				If xtfvig = dFechai Then
					XMARQ = sglmar
				Else
					XMARQ = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
				End If

				If xtipom = 5 Then
					xtfechavig = xthvig
				Else
					xtfechavig = fechaArg(CDate(xtfvig + CLng(CDate("01-01-1900")))) & "&nbsp;&nbsp;" & xthvig
				End If

				Xnlnk = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
				Xnadj = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

				If xtsadj = "S" Then Xnadj = sgladj
				If xtslnk = "S" Then Xnlnk = sgllnk

				Celdak = ""
				Celdak = Celdak & ofv.GenCelda("", "btbs", "", "", "", "", "", sglprf, "si")
				Celdak = Celdak & ofv.GenCelda("", "ut", "", "", "", "", "", "&nbsp;<font color=#999999 ><u>" & xfromusr & "</u></font>&nbsp;&nbsp;", "si")

				Filak = ofv.GenRow("", "", "", "", "", "", "", Celdak, "si")
				sArma = ofv.GenTabla("", "aaa  ", "", "", "", "0", "0", "-1", "-2", "", "", Filak, "si")

				If xtmcod <> "0" Then
					tablan = tablan & "<tr >"
					tablan = tablan & "<td class=AAA colspan=8 style='line-height:2pt;' >&nbsp;</TD>"
					tablan = tablan & "</TR>"



					tablan = tablan & "<tr bgcolor=" & xTrColor & " >"
					tablan = tablan & "<td class=" & xfclas & " width=1% bgcolor=#FFFFFF >" & sImage & "</TD>"
					tablan = tablan & "<td class=vt width=55%  STYLE='text-align:justify;background-color:#DDDDDD;filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & SESSION("CMFPNTIT") & ", endColorStr=" & xTrColor & ", gradientType=1);' >&nbsp;&nbsp;<a class=tt >" & xtfechavig & "</a></TD>"
					tablan = tablan & "<td class=ut width=13% >&nbsp;</TD>"


					If SESSION("NOVFULL") = "S" Then

						tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
						tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
						tablan = tablan & "<td class=btba width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnadj & "</TD>"
						tablan = tablan & "<td class=btbn width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnlnk & "</TD>"
						tablan = tablan & "<td class=qt width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & XMARQ & "</TD>"
					Else
						tablan = tablan & "<td class=UT COLSPAN=3 style='border-style:solid;border-width:1;border-color:#eeeeee;' >&nbsp;</TD>"
						tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
						tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
					End If


					tablan = tablan & "</TR>"

				End If

				tablan = tablan & "<tr bgcolor=" & xTrColor & " >"
				tablan = tablan & "<td class=uti colspan=8 STYLE='text-align:justify;line-height:30px;' ><font color=#999999 >" & xtmsgimg & "&nbsp;" & xtmsglnk & "</font></TD>"
				tablan = tablan & "</TR>"



				If SESSION("NOVFULL") = "S" Then

					If xtmcod <> "0" Then


						tablan = tablan & "<tr bgcolor=" & xTrColor & " height=20 >"
						tablan = tablan & "<td class=ut colspan=8 valign=top align=right style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' >" & sArma & "</TD>"
						tablan = tablan & "</TR>"

					End If

				Else

					If xtmcod <> "0" Then


						tablan = tablan & "<tr bgcolor=" & xTrColor & " height=20 >"
						tablan = tablan & "<td class=ut colspan=8 valign=top align=right  >&nbsp;</TD>"
						tablan = tablan & "</TR>"

					End If

				End If


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If

			Next


		End If



		tablan = tablan & "</table>"
		tablan = tablan & "</DIV>"
		tablan = tablan & "</div>"



	End If

    if  session("APLMDCLI") <> "BGAL" then
	 SFAVORITOS = Favoritos("")
	 SSUSCRIPCION = SUSCRIPCION("")
	 SAUTORIZACION=STRAUTORIZACION
	end if






	shtml = shtml & "<div style='background-color:#ffffff; margin: 0 auto; padding: 0px 0 0px; font: 0.85em arial;'>"

	If session("APLMDCLI") ="BGAL" Then
		shtml = shtml & " <ul class=""tabs"" STYLE='DISPLAY:NONE'>"
	Else
		shtml = shtml & " <ul class=""tabs"" >"
	End If




	If stab = "view1" Then

		shtml = shtml & " <li  class=""selected"" ><a class='nt'  href=""#"" rel=""view1"">NOVEDADES</a></li>"
	Else
		shtml = shtml & " <li   ><a class='nt'  href=""#"" rel=""view1"">NOVEDADES</a></li>"
	End If

	shtml = shtml & " <li><a class='nt'  href=""#"" >|</a></li>"


	If stab = "view2" Then

		shtml = shtml & " <li class=""selected"" ><a class='nt' href=""#"" rel=""view2"">MIS FAVORITOS</a></li>"
	Else
		shtml = shtml & " <li ><a class='nt' href=""#"" rel=""view2"">MIS FAVORITOS</a></li>"
	End If

	shtml = shtml & " <li><a class='nt'  href=""#"" >|</a></li>"

	If stab = "view3" Then
		shtml = shtml & " <li  class=""selected"" ><a id=Xsus class='nt'  href=""#"" rel=""view3"">MIS SUSCRIPCIONES</a></li>"

	Else

		shtml = shtml & " <li   ><a class='nt' id=Xsus  href=""#"" rel=""view3"">MIS SUSCRIPCIONES</a></li>"

	End If
	shtml = shtml & " <li><a class='nt'  href=""#"" >|</a></li>"

	If stab = "view4" Then
		shtml = shtml & " <li  class=""selected"" ><a id=Xsus class='nt'  href=""#"" rel=""view4"">AUTORIZACIONES</a></li>"

	Else

		shtml = shtml & " <li   ><a class='nt' id=Xsus  href=""#"" rel=""view4"">AUTORIZACIONES</a></li>"

	End If

	shtml = shtml & " <li   ><a class='nt'  href=""#"" >|</a></li>"

	If stab = "view5" Then
		shtml = shtml & " <li  class=""selected"" ><a id=Xsus class='nt'  href=""#"" rel=""view5"">DOCUMENTOS A VENCER</a></li>"

	Else

		shtml = shtml & " <li   ><a class='nt' id=Xsus  href=""#"" rel=""view5"">DOCUMENTOS A VENCER</a></li>"

	End If

	shtml = shtml & "</ul>"







	If colsumi > colsumd Then
		ctotspanall = colsumi
	Else

		ctotspanall = colsumd
	End If

	cfilltot = 0

	'ARMA EL MENU PRINCIPAL PARA GALICIA
	Select Case session("APLMDCLI")
		Case "BGAL"


         ' shtml = shtml & "</DIV >"
         ' shtml = shtml & "</DIV >"


          select case stab

          case "view1"

            '  If SESSION("CORNOVEFIMG1") = "NO" Then
			 '		  			XTABLAN2 = "<tr ><td class=isttBF width=50% valign='top'  style='xborder-left-style:solid;xborder-left-width:1;xborder-color:#dddddd;' >"
			 '		  			XTABLAN2=XTABLAN2   & "</td>"
			 '		  		Else
			 		  			'"background-image: url(" & chr(34) & "imagenes/NQN-lago2.jpg" & chr(34) & ") "
			 		  			XTABLAN2 = "<tr ><td class=isttBF width=50% valign='top'  xstyle='xborder-left-style:solid;xborder-left-width:1;xborder-color:#dddddd;background-image: url(" & chr(34) & "IMGNOVEDADES/FONDOS/" & SESSION("CORNOVEFIMG1") & chr(34) & ")' >"
			 		  			XTABLAN2=XTABLAN2 & tablabFORM & "</td>"

			 	'	  		End If


			 		  '		If SESSION("CORNOVEFIMG2") = "NO" Then
			 		  '			XTABLAN21 = "<td class=isttBF width=50% valign='top'       style='xborder-left-style:solid;xborder-left-width:1;xborder-color:#dddddd;' >"
			 		  	'		XTABLAN21 =XTABLAN21 &  "</td>"

			 		  	'	Else
			 		  			XTABLAN21 = "<td class=isttBF width=50% valign='top'       Xstyle='xborder-left-style:solid;xborder-left-width:1;xborder-color:#dddddd;background-image: url(" & chr(34) & "IMGNOVEDADES/FONDOS/" & SESSION("CORNOVEFIMG2") & chr(34) & ")' >"
			 		  			XTABLAN21 =XTABLAN21 & tablabcc & "</td>"
			 		      '  End If

                            fila="<tr>" &  XTABLAN2 & XTABLAN21 & "</tr>"
                            'Celda= "<td VALIGN=TOP>"& ULTFORMCONS("S")& "</td>"
                             Celda= "<td VALIGN=TOP>"
							TNTabla = 0
							INTabla = 0
							xnohay1 = "S"
							xnohay2 = "S"
							saleSel = "N"
							saleSel = "N"
							sBase="K"
							if   atit(11,2)="S" then
							 TNTabla = noveselec(11)

							 UltimosDocConsultados=noveformatEX(11, Menu,"S","")
							else
							 UltimosDocConsultados=""
							end if
							sbase=""
							celda=celda &  UltimosDocConsultados

                            celda=CELDA  & "</td>"


                            Celda=CELDA  & "<td VALIGN=TOP>"
                            TNTabla = 0
							INTabla = 0
							xnohay1 = "S"
							xnohay2 = "S"
							saleSel = "N"
                            saleSel = "N"
                            if   atit(5,2)="S" then
							 TNTabla = noveselec(5)
                             UltimosDocConsultados=noveformatEX(5, Menu,"S","")
                            else
                             UltimosDocConsultados=""
                            end if

                            celda=celda &  UltimosDocConsultados

                            celda=CELDA  & "</td>"

                            fila=fila & "<tr>" &  celda & "</tr>"
                            Celda= "<td  style=""padding-bottom:10px"" align=center>"& LINKBSQ2 & "</td>"
                            Celda= celda & "<td  style=""padding-bottom:10px"" align=center>" & LINKBSQ2BG &"</td>"
                            fila=fila & "<tr >" &  celda & "</tr>"


                            Celda= "<td  CLASS=isttBF height=50px colspan=2 dstyle='text-align:left' ><div dclass=ISTTBF>Últimas Publicaciones</div></td>"
                            fila=fila & "<tr>" &  celda & "</tr>"
                            TNTabla = 0
							INTabla = 0
							xnohay1 = "S"
							xnohay2 = "S"
							saleSel = "N"
							dim xtipoT
                            xtipoT="K"
							TNTabla = noveselec(3)

							UltimasPublicaciones=noveformatEX(3, Menu,"N","")


                            Celda= "<td VALIGN=TOP>" & UltimasPublicaciones & "</td>"

                             TNTabla = 0
							INTabla = 0
							xnohay1 = "S"
							xnohay2 = "S"
							saleSel = "N"
							xtipoT="D"
							TNTabla = noveselec(3)

						    UltimasPublicaciones=noveformatEX(3, Menu,"N","")


                            Celda=celda &  "<td VALIGN=TOP>" & UltimasPublicaciones & "</td>"



                            fila=fila & "<tr>" &  celda & "</tr>"
                            StitN=atit(2,1)
                            Celda= "<td  CLASS=isttBF height=50px colspan=2 >" & StitN & "</td>"
							fila=fila & "<tr>" &  celda & "</tr>"

                            TNTabla = 0
							INTabla = 0
							xnohay1 = "S"
							xnohay2 = "S"
							saleSel = "N"

							TNTabla = noveselec(2)
							Tutoriales=noveformatEX(2, Menu,"N","")


                            Celda= "<td colspan=2 >" & Tutoriales & "</td>"
                            fila=fila & "<tr>" &  celda & "</tr>"
			 		        tablan = "<table width='100%' xborder xcellspacing=2 xcellpadding=2  order xstyle='border-style:solid;border-width:1;border-color:#dddddd;' >"
			 			    tablan = tablan & fila
			 		        tablan = tablan & "</table>"
			           shtml =shtml & tablan

         ' shtml = shtml & "</DIV>"
          case "view2"
         '  shtml = shtml & "</div><div id=""view2"" class=""tabcontent"">"
		   shtml = shtml & Favoritos("N")
		 '  shtml = shtml & "      </div>"
          case "view3"
         '  shtml = shtml & " </div><div id=""view3"" class=""tabcontent"">"
		   shtml = shtml & SUSCRIPCION("N")
		 '  shtml = shtml & "</div>"
          case "view4"
          ' shtml = shtml & " </div><div id=""view4"" class=""tabcontent"">"

                TNTabla = 0
				INTabla = 0
				xnohay1 = "S"
				xnohay2 = "S"
				saleSel = "N"
                TNTabla = noveselec(4)
           'TNTabla=INTabla
		   shtml = shtml   & noveformatEX(4, Menu,"N","")
		 '  shtml = shtml & "     </div>"
          case "view5"
           'shtml = shtml & " </div><div id=""view5"" class=""tabcontent"">"
           TNTabla = 0
		   INTabla = 0
		   xnohay1 = "S"
		   xnohay2 = "S"
		   saleSel = "N"
           xxxx = noveselec(7)
		   shtml = shtml & noveformatEX(7, Menu,"N","")
		  ' shtml = shtml & "     </div>"

		  end select
         ' shtml = shtml & "</DIV >"

		Case Else ' ARMA 3 COLUMNAS PARA EL RESTO DE LOS CLIENTES

        shtml = shtml & "<div  class=""tabcontents"" >"
	    sHTML =shtml&   "<div id=""view1"" class=""tabcontent"">"
	    shtml = shtml & "<DIV id='DTGENERAL' xSTYLE='display: table;' >"

	    Sfpath=""
	    IF Session("EMPMUN") = "S" THEN Sfpath="Emp" & session("un") & "/"

		If SESSION("CORNOVEFIMG1") = "NO" Then
			XTABLAN2 = "<tr ><td class=ut width=30% valign='top'  style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
		Else
			'"background-image: url(" & chr(34) & "imagenes/NQN-lago2.jpg" & chr(34) & ") "
			XTABLAN2 = "<tr ><td class=ut width=30% valign='top'  style='border-left-style:solid;border-left-width:1;border-color:#dddddd;background-image: url(" & chr(34) & Sfpath & "IMGNOVEDADES/FONDOS/" & SESSION("CORNOVEFIMG1") & chr(34) & ")' >"
		End If


		If SESSION("CORNOVEFIMG2") = "NO" Then
			XTABLAN21 = "<td class=ut width=35% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
		Else
			XTABLAN21 = "<td class=ut width=35% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;background-image: url(" & chr(34) & Sfpath & "IMGNOVEDADES/FONDOS/" & SESSION("CORNOVEFIMG2") & chr(34) & ")' >"
		End If

		If SESSION("CORNOVEFIMG3") = "NO" Then
			XTABLAN22 = "<td class=ut width=35% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
		Else
			XTABLAN22 = "<td class=ut width=35% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;background-image: url(" & chr(34) & Sfpath & "IMGNOVEDADES/FONDOS/" & SESSION("CORNOVEFIMG3") & chr(34) & ")' >"
		End If

		'response.write ttitsecc
		For ii = 1 To ttitsecc

			cworka = CInt(ii / 2)
			cworka = cworka * 2

			If cworka = ii Then
				colsumd = colsumd + atitsecciones(ii, 2)
			Else
				colsumi = colsumi + atitsecciones(ii, 2)
			End If

			xrowspan1 = " rowspan=" & atitsecciones(ii, 2)
			xrowspan1 = " "

			ctotspani = ctotspani + atitsecciones(ii, 2)

			If ii = 3 Or ii = 5 Then
				XTABLAN22 = XTABLAN22 & atitsecciones(ii, 3)
			Else
				XTABLAN21 = XTABLAN21 & atitsecciones(ii, 3)
			End If

			ii = ii + 1
			If ii > ttitsecc Then
				xcolworka = "&nbsp;"
				xrowspan2 = ""
			Else
				xcolworka = atitsecciones(ii, 3)
				xrowspan2 = " rowspan=" & atitsecciones(ii, 2)
				ctotspand = ctotspand + atitsecciones(ii, 2)
				xrowspan2 = ""
				' comentado 31/7/2018
				'XTABLAN22 = XTABLAN22 & xcolworka
			End If

			XTABLAN22 = XTABLAN22 & xcolworka

		Next



		If session("CORMODUSOLABM")<> "S"  Then
			tablaLN=ultimosDocConsultados & tablaLN
		Else
			XTABLAN22= XTABLAN22 & ultimosDocConsultados
		End If

		XTABLAN2 = XTABLAN2 & tablab &  tabla1 & tablabFORM & tabla1f & tablaLN & "</Td>" & XTABLAN21 & "</Td>" & XTABLAN22 & "</Td>" & "</TR>"


		tablan = "<table width='100%' xcellspacing=2 xcellpadding=2  xstyle='border-style:solid;border-width:1;border-color:#dddddd;' >"
		tablan = tablan & XTABLAN2
		tablan = tablan & "</table>"

		shtml = shtml & tablan
		shtml = shtml & "</DIV >"
		shtml = shtml & "</div><div id=""view2"" class=""tabcontent"">"
		shtml = shtml & SFAVORITOS

		shtml = shtml & "      </div>"
		shtml = shtml & " <div id=""view3"" class=""tabcontent"">"
		shtml = shtml & SSUSCRIPCION
		shtml = shtml & "     </div>"
		shtml = shtml & " <div id=""view4"" class=""tabcontent"">"
		shtml = shtml & STRAUTORIZACION
		shtml = shtml & "     </div>"
		shtml = shtml & " <div id=""view5"" class=""tabcontent"">"
		shtml = shtml & STRDOCVENCER
		shtml = shtml & "     </div>"

	    shtml = shtml & " </div>"
	End Select






	Agregar = ""
	Buscar = ""
	CELDA = ""

	shtml = shtml & ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "no", Agregar, Buscar, CELDA)

	shtml = shtml & WSPDTTSCP


	'' *****  JAVASCRIPT PARA VISTA PRELIMINAR DE NOVEDADES ******
	wstyl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"
	wstyl = wstyl & "border-style:solid;border-width:1;border-color:#000000;background-color:#ffffff;"

	shtml = shtml & "<div id='mt2view'  style='Z-INDEX:999999;display:none;position:absolute;top:200;left:300;height:30%;width:50%;overflow:auto;" & wstyl & "' >"

	shtml = shtml & "<table width=99% >"
	shtml = shtml & "<tr height=10 >"
	shtml = shtml & "<td id='mt2viewl1c1' class='tt'  width=90% style='border-style:solid;border-width:1;border-color:#000000;padding-left:5px;background-color:" & SESSION("CMFMENU") & "' >Mostrando vista de </td>"
	shtml = shtml & "<td id='mt2viewl1c2'  class='tt'  align=center style='border-style:solid;border-width:1;border-color:#000000;padding-left:5px;background-color:" & SESSION("CMFMENU") & "'> <a class=nt href='' onclick=" & Chr(34) & "document.all.mt2view.style.display='none'; return false;" & Chr(34) & " style=""cursor:pointer;border-style:outset;border-width:1;border-color:#000000;background-color:#eeeeee;"" ><font size=2>Cerrar</font></a></td>"
	shtml = shtml & "</tr>"

	shtml = shtml & "<tr height=100% >"
	shtml = shtml & "<td  id='DX4' colspan=2  class='uT' style='background-color:#ffffff;' >contenido</td>"
	shtml = shtml & "</tr>"
	shtml = shtml & "</table>"

	shtml = shtml & "</div >"

	shtml = shtml & "<script language=javascript > "
	shtml = shtml & "function WKlevanta(wkvista,e) { " & Chr(13)
	shtml = shtml & "  document.all.mt2viewl1c1.innerHTML='<h5> </h5>';var xxp=getPosition(e) -200;" & Chr(13)
	shtml = shtml & " var frd= document.getElementById('mt2view') ;frd.style.top=xxp  + 'px';" & chr(13)

	shtml = shtml & "  document.all.mt2view.style.display='block';   " & Chr(13)
	shtml = shtml & " } " & Chr(13)
	shtml = shtml & " </script> "
	'' *****  JAVASCRIPT PARA VISTA PRELIMINAR DE NOVEDADES ******

    Call ofv.cerrarconn(cn)
	Call ofv.cerrarconn(cn0)
    stab=request.querystring("tab")
	if stab="" and session("APLMDCLI")="BGAL" then
	 WSPFORME = "<FORM ID='fnoved' NAME='fnoved' ACTION='../coress20/coriconos.asp' TARGET='iconos' style='display:none;' METHOD='POST' ></FORM>"
	 shtml = WSPFORME & shtml & "<script language=javascript > "
	 shtml = shtml & "document.addEventListener('DOMContentLoaded', function() {"
	 shtml = shtml &    "var fr= document.getElementById('fnoved');fr.submit();"
     shtml = shtml & "})"
     shtml = shtml & " </script> "
    end if
	shtml = shtml &  "</body></html>"



	response.write shtml
	response.write xxform
	shtml = "<script type='text/javascript'>"
	shtml = shtml & " var scw= screen.availWidth;" & Chr(13)
	'shtml = shtml & " if (scw <= 800){document.Subconsulta.textot.size=20;} else {document.Subconsulta.textot.size=40;}" & Chr(13)
	shtml = shtml & "function disp_confirm(disp_msg,disp_frm)"
	shtml = shtml & "  {"
	shtml = shtml & "  var r=confirm(disp_msg);"
	shtml = shtml & "  if (r==true)"
	shtml = shtml & "    {"
	shtml = shtml & "  var fr= document.getElementById(disp_frm);"
	'sHTML = sHTML & "    alert(fr.name);"
	shtml = shtml & "    fr.submit();"
	shtml = shtml & "    }"
	shtml = shtml & "  else"
	shtml = shtml & "    {"
	shtml = shtml & "    alert('Se cancelo la accion!');"
	shtml = shtml & "    }"


	shtml = shtml & "  }"



	shtml = shtml & "</script>"
	response.write shtml

End If

Function EsqnomP(xtipo, xcodigo, ofv)

	Set xcn = ofv.conectar(ofv.strconn2)


	If xtipo = "RSK" Then

		StrSql = "SELECT * FROM  ORRSKCIRCUITO "
		StrSql = StrSql & " Where OBJID = 0" & xcodigo
		Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
		If Not xrs.EOF Then
			If xrs("APLIC") = "OR" Then
				xdescri = "CIRCUITO AUTOEVALUACION"
			Else
				xdescri = "PLAN DE CONTINGENCIAS"
			End If

			StrSql = "Select * From objetos where codobjeto = 0" & xrs("CODOBJETO")
			Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
			If Not xrs.EOF Then
				xiden = xrs(1)
				xobjdes = xrs(3)
			End If


		End If
		Call ofv.cerrarconsulta(xrs)

		'''''sTipobjetoX = Capitalize(xdescri) & ": " & Capitalize(xobjdes)
		sTipobjetoX = xdescri & ": " & xobjdes

	Else


		sTipobjetoX = xcodigo
		sql="select * from tipobjeto where tipobjeto='" & xtipo &"'"
		'  RESPONSE.WRITE SQL
		Set RSM = ofv.crearconsultaEx(sql, xcn, 1, parametros)

		tipocpt=""
		If Not rsm.eof Then tipocpt=rsm("tipocpt")
		Call OFV.CERRARCONSULTA(RSM)

		Select Case tipocpt


			Case "PD"
			StrSql2 = "Select a.* ,b.nombre From PEDIDOS a "
			StrSql2 = StrSql2 & " inner join pedidosdet b on (b.idpedido=a.objid) "
			StrSql2 = StrSql2  & " Where a.OBJID = 0" & xcodigo
			' response.write strsql2
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			If Not RSM.EOF Then

				XX=GETINFOUSER(RSM("SOLICITANTE"))
				sTipobjetoX= " Pedido Nro: " & RSM(0) & " De "  & SESSION("SGU_NOMBRE")

				sTipobjetoX =  sTipobjetoX &  "&nbsp;&nbsp;&nbsp;" & Left(RSM("nombre") ,100) & "..."


			End If
			Call OFV.CERRARCONSULTA(RSM)

			Case "SN"
			StrSql2 = "Select a.* ,b.nombre From PEDIDOS a "
			StrSql2 = StrSql2 & " inner join pedidosdet b on (b.idpedido=a.objid) "
			StrSql2 = StrSql2  & " Where a.OBJID = 0" & xcodigo
			' response.write strsql2
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			If Not RSM.EOF Then

				XX=GETINFOUSER(RSM("SOLICITANTE"))
				sTipobjetoX= " Pedido Nro: " & RSM(0) & " De "  & SESSION("SGU_NOMBRE")

				sTipobjetoX =  sTipobjetoX &  "&nbsp;&nbsp;&nbsp;" & Left(RSM("nombre") ,100) & "..."


			End If
			Call OFV.CERRARCONSULTA(RSM)

			Case Else

			StrSql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
			Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
			If Not RSM.EOF Then




				If SESSION("APLMDCLI") = "BC" Then
					sTipobjeto1 = (RSM(1)) & ":&nbsp;"
					sTipobjeto2 = RSM(3) & ""
					sTipobjeto3 = " - " & (RSM(2) & "")
					sTipobjeto4 = " (" & (RSM(5) & "") & " "
					sTipobjeto5 = (RSM(6) & "") & ") "
				Else
					'''''sTipobjeto1 = Capitalize(RSM(1)) & ":&nbsp;"
					sTipobjeto1 = RSM(1) & ":&nbsp;"
					'''''sTipobjeto2 = Capitalize(RSM(3) & "")
					sTipobjeto2 = RSM(3) & ""

					'''''sTipobjeto3 = " - " & Capitalize(RSM(2) & "")
					sTipobjeto3 = " - " & RSM(2) & ""
					'''''sTipobjeto4 = " (" & Capitalize(RSM(5) & "")
					sTipobjeto4 = " (" & RSM(5) & " )"
					If UCase(RSM(5) & "") = "O" Then
						sTipobjeto5 = ")"
					Else
						'''''sTipobjeto5 =  " " & Capitalize(RSM(6) & "") & ") "
						sTipobjeto5 = " " & RSM(6) & "" & ") "
					End If
				End If
				If SESSION("FULLDOCN") = "1" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2
				ElseIf SESSION("FULLDOCN") = "2" Then
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
				Else
					sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
				End If



			End If
		End Select

	End If





	Call ofv.cerrarconn(xcn)


	EsqnomP = sTipobjetoX


End Function


Function fechaArg(f)
	dia = Day(f)
	mes = Month(f)
	ano = Year(f)
	fechaArg = dia & "/" & mes & "/" & ano
End Function

Function wmnbuscamenu(wmnSubmenu,dato)

	Set cx0 = ofv.conectar(ofv.strconn0)
	Set cx1 = ofv.conectar(ofv.strconn1)

	xsel = 1
	StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & SESSION("CorApli") & "' "

	Set rs = ofv.crearconsultaEx(StrSql, cx1, 1, "")

	If Dato="N" Then

		If Not rs.eof Then
			wmnbuscamenu="OK"
		Else
			wmnbuscamenu="NO"
		End If
		'response.write  wmnbuscamenu
	Else
		If Not rs.EOF Then
			If rs.fields.Count > 6 Then
				xsel = 6
				Call ofv.cerrarconsulta(rs)
				StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & SESSION("CorApli") & "' order by Morden "

				Set rs = ofv.crearconsultaEx(StrSql, cx1, 1, "")
			Else
				xsel = 1
			End If
		End If
		Do Until rs.EOF



			atrok = "N"
			StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'M' "
			StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
			StrSql = StrSql & " And usuario = '" & SESSION("segperfil") & "' And "
			StrSql = StrSql & "  Aplicacion = '" & SESSION("CorApli") & "'"

			Set rs1 = ofv.crearconsultaEx(StrSql, cx0, 1, "")
			If rs1.EOF Then
				atrok = "O"
			ElseIf rs1(1) Then
				atrok = "S"
			Else
				atrok = "N"
			End If
			Call ofv.cerrarconsulta(rs1)

			xcount = 0

			If atrok = "O" Then

				StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'M' "
				StrSql = StrSql & " And alias = '" & rs(xsel) & "' And  "

				If SESSION("DELLOG") = "S" Then

					XMFTGRP = Split(SESSION("SegGrpLst"), Chr(9))
					xcount = UBound(XMFTGRP) + 1
					XMFSEP = ""
					StrSql = StrSql & " grupo in ("

					For XMFI = 0 To UBound(XMFTGRP)
						StrSql = StrSql & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
						XMFSEP = ","
					Next

					StrSql = StrSql & ") "

				Else
					StrSql = StrSql & " grupo = '" & SESSION("seggrp") & "' "
				End If


				StrSql = StrSql & " And Aplicacion = '" & SESSION("CorApli") & "'"
				Set rs2 = ofv.crearconsultaEx(StrSql, cx0, 1, "")
				If rs2.EOF Then
					atrok = "O"
				Else
					If SESSION("DELLOG") = "S" Then
						XMFSUM = 0
						XMFSUMN = 0
						Do Until rs2.EOF
							XMFSUM = XMFSUM + 1
							If rs2(1) Then
								atrok = "S"
								Exit Do
							Else
								atrok = "N"
								XMFSUMN = XMFSUMN + 1
							End If
							rs2.movenext
						Loop

						If atrok = "N" Then
							If xcount > XMFSUMN Then atrok = "S"
						End If

					Else
						If rs2(1) Then
							atrok = "S"
						Else
							atrok = "N"
						End If
					End If
				End If
				Call ofv.cerrarconsulta(rs2)

			End If


			If atrok = "O" Or atrok = "S" Then


				sInput = Trim(rs(xsel))
				sHref = Trim(rs(2))
				sTarget = "info"
				sPath = SESSION("CorPath")
				If rs(3) <> "" Then atrok = "N"
				spt = InStr(rs(2), ".asp")
				xesm = "N"
				If spt > 1 Then
					If Left(rs(2), spt - 1) = "Submenus" Then atrok = "N"
				End If
				sseg = InStr(sHref, "?")
				If sseg = 0 Then
					sHref = sHref & "?"
				Else
					sHref = sHref & "&"
				End If
				sHref = sHref & "seg=" & ofv.convertircar(sInput, " ", "_") & "&MNN=" & ofv.convertircar(UCase(rs(1)), " ", "_")

				If atrok <> "N" Then


					ii = ii + 1

					atit(ii,1) = UCase(rs(1))
					awmnasp(ii) = sPath & sHref


				End If





			End If




			rs.movenext
		Loop

	End If
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cx1)
	Call ofv.cerrarconn(cx0)




End Function

Function noveselec(xmenu)

	'' xscolor = "#999999 size=2 "
	'response.write xmenu & "  "
	xscolor = " class=istt "    '"#0000 style='font-size:8pt;font-weight:normal;' "

	Do Until saleSel = "S"

		Select Case xmenu

			Case 1


			Case 2

			'''    session("ModMain") = "CL"

			StrSql = "SELECT distinct * FROM csnovedades "
			StrSql = StrSql & " Where fecexp > " & dFechai
			StrSql = StrSql & " and fecinicio < " & dFechai + 1
			StrSql = StrSql & " And ("
			StrSql = StrSql & " tipo = 1 or tipo = 2 or tipo = 51"
			StrSql = StrSql & ")"
			StrSql = StrSql & " And ("
			StrSql = StrSql & " (usugrp = '" & SESSION("usuario") & "') or "
			StrSql = StrSql & " (usugrp = 'TODOS' and sector = 999 )    or "
			StrSql = StrSql & " (usugrp = 'TODOS' and sector = 0" & SESSION("UN") & ") OR "
			StrSql = StrSql & " (usugrp = 'CORP' and sector = 0" & SESSION("UN") & ")"
			StrSql = StrSql & ") "
			StrSql = StrSql & " order by fecinicio desc, tipo desc, importancia, objid desc"
			'RESPONSE.WRITE STRSQL
			Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If rsx.EOF Then
				xnohay1 = "S"
			Else
				xnohay1 = "N"
				Do Until rsx.EOF
					INTabla = INTabla + 1

					xfromusr = rsx(11)

					xxxx = getInfoUser(rsx(11))

					xfromusr = SESSION("SGU_nombre")

					ANTabla(INTabla, 0) = rsx(1)
					ANTabla(INTabla, 1) = CLng(rsx(2))
					ANTabla(INTabla, 2) = CLng(rsx(6))
					ANTabla(INTabla, 3) = CLng(rsx(3))
					ANTabla(INTabla, 4) = rsx(12)
					ANTabla(INTabla, 5) = "N"
					ANTabla(INTabla, 6) = "N"
					ANTabla(INTabla, 7) = xfromusr
					ANTabla(INTabla, 8) = rsx(5)
					ANTabla(INTabla, 9) = "&nbsp;"
					ANTabla(INTabla, 10) = "XXMSGXX"
					ANTabla(INTabla, 11) = rsx(2)
					ANTabla(INTabla, 12) = rsx(13)
					'''ANTabla(INTabla, 13) = Capitalize(rsx(14))
					'''ANTabla(INTabla, 14) = Capitalize(rsx(15))

					If (rsx("usugrp") = "CORP") And (SESSION("EMPMUN") = "S") Then
						smasc = "&nbsp;&nbsp;&nbsp;<font color=#993333 ><font class=btic title='Corporativo' >G</font></font>"
					Else
						smasc = "&nbsp;"
					End If

					If SESSION("APLMDCLI") = "BGAL" And rsx(2) & "" = "1" Then
						ANTabla(INTabla, 13) = rsx(14) & smasc
					Else
						''''ANTabla(INTabla, 13) = Capitalize(rsx(14))
						ANTabla(INTabla, 13) = rsx(14) & smasc
					End If
					If SESSION("APLMDCLI") = "BGAL" And rsx(2) & "" = "1" Then
						ANTabla(INTabla, 14) = rsx(15)
					Else
						''''ANTabla(INTabla, 14) = Capitalize(rsx(15))
						If SESSION("APLMDCLI") = "GP" Then
							ANTabla(INTabla, 14) = rsx(15) & "<BR><BR><BR>" & RSX(5)
						Else
							ANTabla(INTabla, 14) = rsx(15)
						End If
					End If

					ANTabla(INTabla, 15) = rsx(16)
					ANTabla(INTabla, 16) = rsx(18)

					ANTabla(INTabla, 17) = rsx(19) & ""
					ANTabla(INTabla, 18) = rsx(20) & ""
					ANTabla(INTabla, 19) = rsx(21) & ""
					ANTabla(INTabla, 20) = rsx(22) & ""

					ANTabla(INTabla, 21) = "0"



					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)

			Case 3, 9

			XEMPRE = Split(session("EmpHbl"),Chr(9))

			DELIM=""
			idempre = ""

			For i=0 To UBound(XEMPRE)

				idempre = idempre & DELIM & XEMPRE(i)
				DELIM = ","

			Next

			'''    session("ModMain") = "CL"
			StrSql = "SELECT distinct a.*, isnull(b.cliente,1) FROM csnovedades a "
			select case xtipoT
			  case "K"
			  StrSql = StrSql & "inner join tipobjeto a1 on (a.tipobjeto=a1.tipobjeto and a1.TIPOCPT = 'K') "
			  case "D"
			  StrSql = StrSql & "inner join tipobjeto a1 on (a.tipobjeto=a1.tipobjeto and a1.TIPOCPT = 'D') "

			end select
            '--------------------------------------------------------------------------
			'NUEVO 02/08/2021 GP --> ALE
			'10317-0004/20
			SELECT CASE session("APLMDCLI")

			CASE "GP"

			   StrSql = StrSql & " inner JOIN objetos c ON (c.codobjeto = a.codobjeto AND c.TIPOBJETO=a.TIPOBJETO)"
               StrSql = StrSql & " inner JOIN OBJMEMPALC  c1 ON (c1.codobjeto = a.codobjeto AND c1 .TIPOBJETO=a.TIPOBJETO)"

			END SELECT
           '--------------------------------------------------------------------------
			StrSql = StrSql & " left join clienteplan b on (a.sector=b.id) "
			StrSql = StrSql & " Where a.fecexp > " & dFechai
			StrSql = StrSql & " and a.fecinicio <  " & dFechai + 1
			StrSql = StrSql & " And (a.tipo = 3 or a.tipo = 9)"
			'If session("APLMDCLI")<> "BPN" Then			'Se quita excepción para BPN
				StrSql = StrSql & " And ("
				StrSql = StrSql & "    (a.usugrp = '" & session("usuario") & "') "
				StrSql = StrSql & " or ( a.usugrp = 'TODOS' and a.sector = 999)"
				If session("UN") = 1 Then
					StrSql = StrSql & " or ( a.usugrp = 'TODOS' and a.sector in (" & idempre & "))"
				Else
					StrSql = StrSql & " or ( a.usugrp = 'TODOS' and a.sector = 0" & session("UN") & ")"
				End If
				SELECT CASE session("APLMDCLI")

			     CASE "GP"
				StrSql = StrSql & " or (c.objun=1 and c1.codempre=" & session("UN") & ")"
				END SELECT
				StrSql = StrSql & "     ) "
			'Else
				'  StrSql = StrSql & " And ("
				'    StrSql = StrSql & "    (usugrp = '" & session("usuario") & "') "
				'  StrSql = StrSql & "  ( usugrp = 'TODOS' and sector = 999)"
				' StrSql = StrSql & " or ( usugrp = 'TODOS' and sector = 0" & session("UN") & ")"
				'   StrSql = StrSql & " or ( usugrp = 'TODOS' and sector = 0)"
				' StrSql = StrSql & "     ) "
			'End If

			StrSql = StrSql & " order by a.fecinicio desc, a.importancia, objid desc"

			Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If rsx.EOF Then
				xnohay1 = "S"

			Else
				contador = 0
				xvinoant = "no"
				Do Until rsx.EOF

					contador = contador + 1

					If xvinoant <> rsx(1) Then
						xvinoant = rsx(1)


						xvale = "N"
						xcorpo="N"
						StrSql = "SELECT * FROM objetos "
						StrSql = StrSql & " Where codobjeto = 0" & rsx(8)
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							If rsxp(7) & "" ="1" Then
								xcorpo="S"
							Else
								If session("UN") <> 1  Then
									If rsxp(7) & "" <> session("UN") & "" Then xvale = "N"
								End If
							End If

						End If



						Call ofv.cerrarconsulta(rsxp)
						'response.write atit(3,3) & "<br>"
						If xvale = "S" And contador <= atit(3,3) Then

							xnohay1 = "N"
							INTabla = INTabla + 1


							xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


							sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(8) & "&Tipo=" & rsx(9)
							sActcon = sActcon & "&Datos=*smenu=zzzzz*lista=3"
							sHref = "&nbsp;"
							' sActcon="http://www.google.com"
							' response.write sActcon
							sFuncion = " onclick=" & Chr(34) & "AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)


							wmsgx = Trim(rsx(5) & "")

							If Len(wmsgx) < 5 Then
								wmsgx = EsqnomPx(rsx(9), rsx(8), ofv) '& " VIGENTE A PARTIR DEL " & fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900")))) & ". "
							Else
								'if xmenu=q3 then
								' wmsgx = UCase(wmsgx) ' & "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
								' else'response.write "aa" & len(wmsgx) & rsx(5)
								wmsgx = UCase(wmsgx)  '& "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
								' end if
							End If

							wNomL = wmsgx & "&nbsp;" '& Len(wmsgx)



							sImagen = wNomL

							If xarc = "NO" Then
								xarc = "&nbsp;"
							End If

							sLink = sImagen
							'   XTSLNKM = OFV.GenLink("", "cmb title='Consultar Documento...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
							'XTSLNKM = ofv.GenLink("", "istt title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")
						    XTSLNKM = ofv.GenLink("", "nt title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")
							xfromusr = rsx(11)


							xxxx = getInfoUser(rsx(11))

							xfromusr = SESSION("SGU_nombre")

							ANTabla(INTabla, 0) = rsx(1)
							ANTabla(INTabla, 1) = rsx(2)
							ANTabla(INTabla, 2) = CLng(rsx(6))
							ANTabla(INTabla, 3) = CLng(rsx(3))
							ANTabla(INTabla, 4) = rsx(12)
							ANTabla(INTabla, 5) = "S"
							ANTabla(INTabla, 6) = "N"
							ANTabla(INTabla, 7) = xfromusr
							ANTabla(INTabla, 8) = sLink
							ANTabla(INTabla, 9) = xarc
							ANTabla(INTabla, 10) = XTSLNKM
							ANTabla(INTabla, 11) = xmenu

							If (rsx(13) & "") = "2" Then
								ANTabla(INTabla, 12) = rsx(13)
								''''ANTabla(INTabla, 13) = Capitalize(rsx(14))
								ANTabla(INTabla, 13) = rsx(14)
								If rsx(2) = "9" Then ANTabla(INTabla, 13) = ANTabla(INTabla, 13) & " - " & EsqnomPx(rsx(9), rsx(8), ofv)

								''''ANTabla(INTabla, 14) = Capitalize(rsx(15))
								ANTabla(INTabla, 14) = rsx(15)
								ANTabla(INTabla, 15) = rsx(16)
								ANTabla(INTabla, 16) = rsx(18)
							Else

								ANTabla(INTabla, 12) = 2
								If xcorpo ="S"  And (SESSION("EMPMUN") = "S") Then
									smasc = "&nbsp;&nbsp;&nbsp;<font color=#993333 ><font class=btic title='Corporativo' >G</font></font>"
								Else
									smasc = "&nbsp;"
								End If

								ANTabla(INTabla, 13) = EsqnomPx(rsx(9), rsx(8), ofv) & smasc
								ANTabla(INTabla, 14) = sLink
								ANTabla(INTabla, 15) = "0"
								ANTabla(INTabla, 16) = ""
							End If

							ANTabla(INTabla, 17) = rsx(19) & ""
							ANTabla(INTabla, 18) = rsx(20) & ""
							ANTabla(INTabla, 19) = rsx(21) & ""
							ANTabla(INTabla, 20) = rsx(22) & ""

							ANTabla(INTabla, 21) = rsx(8)

						End If
					End If

					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)





			Case 4

			'''     session("ModMain") = "WK"
			SESSION("Segce") = "S"
			SESSION("Segco") = "S"
			cuenta=0

			StrSql = "SELECT * FROM csnovedades "
			StrSql = StrSql & " Where fecexp > " & dFechai
			StrSql = StrSql & " and fecinicio < " & dFechai + 1
			StrSql = StrSql & " And tipo = 4"
			StrSql = StrSql & " And usugrp = '" & SESSION("usuario") & "' order by fecinicio desc, importancia, objid desc"
			'   RESPONSE.WRITE STRSQL
			Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If rsx.EOF  Then
				xnohay1 = "S"
			Else
				xnohay1 = "N"
				Do Until rsx.EOF
					cuenta=cuenta+1

					INTabla = INTabla + 1

					StrSql = "SELECT archivo FROM tipobjeto "
					StrSql = StrSql & " Where tipobjeto = '" & rsx(9) & "'"
					Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs.EOF Then
						xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
					Else
						xarc = "NO"
					End If
					Call ofv.cerrarconsulta(rs)


					SESSION("ASP") = "novedmenu.asp?menu=4"
					SESSION("sEstado") = "PARA APROBAR"
					SESSION("sEstadon") = "2"
					SESSION("Form2") = "apr"

					aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&Perfil=xxxx&Funcion=E"
					aspeditar = CStr(aspeditar)
					sHref = "&nbsp;"
					' response.write sActcon
					sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

					'''''wNomL = "&nbsp;" & Capitalize(rsx(5)) & "&nbsp;"
					wNomL = "&nbsp;" & rsx(5) & "&nbsp;"

					sImagen = wNomL


					If xarc = "NO" Then
						xarc = "&nbsp;"
					End If

					''''  sLink = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=" & xscolor & " >" & sImagen & "</font>", "si")


					sLink = sImagen
					'XTSLNKM = OFV.GenLink("", "cmb ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
					XTSLNKM = ofv.GenLink("", "istt ", sHref, "", "", sFuncion, "XXMSGXX", "si")


					xfromusr = rsx(11)


					xxxx = getInfoUser(rsx(11))
					xfromusr = SESSION("SGU_nombre")

					ANTabla(INTabla, 0) = rsx(1)
					ANTabla(INTabla, 1) = 4
					ANTabla(INTabla, 2) = CLng(rsx(6))
					ANTabla(INTabla, 3) = CLng(rsx(3))
					ANTabla(INTabla, 4) = rsx(12)
					ANTabla(INTabla, 5) = "S"
					ANTabla(INTabla, 6) = "N"
					ANTabla(INTabla, 7) = xfromusr
					ANTabla(INTabla, 8) = sLink
					ANTabla(INTabla, 9) = xarc
					ANTabla(INTabla, 10) = XTSLNKM
					ANTabla(INTabla, 11) = xmenu
					ANTabla(INTabla, 12) = rsx(13)
					ANTabla(INTabla, 13) = rsx(14)
					ANTabla(INTabla, 14) = rsx(15)
					ANTabla(INTabla, 15) = rsx(16)
					ANTabla(INTabla, 16) = rsx(18)

					ANTabla(INTabla, 17) = rsx(19) & ""
					ANTabla(INTabla, 18) = rsx(20) & ""
					ANTabla(INTabla, 19) = rsx(21) & ""
					ANTabla(INTabla, 20) = rsx(22) & ""

					ANTabla(INTabla, 21) = "0"




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

				StrSql = "SELECT tipobjeto FROM tipobjeto "
				Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				If Not rs.EOF Then
					Do While Not rs.EOF
						xtipos = xtipos & separ & "'" & rs(0) & "'"
						separ = ","
						rs.movenext
					Loop
				Else
					xtipos = xtipos & "' '"
				End If
				xtipos = xtipos & ")"

			End If





			SEGWHERE = " WHERE " & xtipos & " AND ( PSCFIN = 'NO'  or  ( fechadif < 0 and finp=1) ) AND ( (RESPASG = '" & session("Usuario") & "') "


			xxxx = getInfoUser(SESSION("Usuario"))
			xusrlog = SESSION("SGU_nombre")
			'RESPONSE.WRITE xusrlog

			If xusrlog = "NO" And SESSION("USRPRV") = "LACAJA" Then
				SESSION("USRPRV") = "I"
				SESSION("Usuario") = "USRCONSULTA"
				Call getInfoUser(SESSION("Usuario"))
				xusrlog = SESSION("SGU_nombre")
				SESSION("USRPRV") = "LACAJA"
			End If

			' xxxx = getInfoUser(session("Usuario"))
			' xusrlog = session("SGU_nombre")


			'  RESPONSE.WRITE xusrlog &" 222222222222222222222222222" & CHR(13)
			If xusrlog <> "NO" Then


				SEGWHERE = SEGWHERE & "OR ( ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & SESSION("Usuario") & "') ) "
				SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

				'''''''''''''''''''''''''''''''''''''''''''
				If SESSION("DELLOG") = "S" Then

					XMFTGRP = Split(SESSION("SegGrpLst"), Chr(9))
					xcount = UBound(XMFTGRP) + 1
					XMFSEP = ""
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

					For XMFI = 0 To UBound(XMFTGRP)
						SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
						XMFSEP = ","
					Next

					SEGWHERE = SEGWHERE & ") ) )"

				Else
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
				End If


				''''''''''''''''''''''''''''''''''''''''''

				'''   IF  RS(21) > 0 THEN


				If CStr(SESSION("SGU_idcostcenter") & "") <> "0" Then
					SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
					SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & SESSION("SGU_idcostcenter") & "' OR SUBCODIGO = '" & SESSION("SGU_idcostcenter") & "' ) "
					SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
					SEGWHERE = SEGWHERE & " ) ) "
				End If


			End If
			If xusrlog <> "NO" Then
				SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "
			Else
				SEGWHERE = SEGWHERE & " ) ORDER BY PSCCOMIENZO "
			End If

			StrSql = "SELECT * FROM CKTVIEWex  " & SEGWHERE

		'	RESPONSE.WRITE session("CLIEMP") & "<BR>" & chr(13)
			Set rsx = ofv.crearconsultaEx(StrSql, cn1, 1, parametros)
			If rsx.EOF Then
				xnohay2 = "S"
			Else

				Do Until rsx.EOF

					xvale = "N"

					If rsx(3) = "RSK" Then

						StrSql = "SELECT * FROM ORRSKCIRCUITO "
						StrSql = StrSql & " Where OBJID = 0" & rsx(4)
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)


					Else



						StrSql = "SELECT a.*, b.cliente FROM objetos a "
						StrSql = StrSql & " inner join clienteplan b on (a.objun=b.id) "
						StrSql = StrSql & " Where a.codobjeto = 0" & rsx(4)
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                       ' response.write strsql & "<br>"
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

					End If

					StrSql = "SELECT * FROM tipobjeto where (tipocpt='PD' or tipocpt='SN') AND tipobjeto='" & rsx ("objtipo") & "'"
				'	response. write strsql & "<br>"
					Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					tipocpt=""
					If Not rsxp.EOF Then
						xvale = "S"
						xMAINAPL = rsxp("APLIC")
						tipocpt=rsxp("tipocpt")
					End If
					Call ofv.cerrarconsulta(rsxp)

					If session("CLIEMP") = 11111 Then

						sql = "select site from usuarios where ID='" & rsx(6) & "'"
						'response.write sql
						Set rsp = ofv.crearconsultaEx(sql, cn, 1, parametros)
						If Not rsp.eof Then
							If rsp(0) & "" = session("UN") & "" Then
								xvale = "S"
							Else
								xvale = "N"
							End If
						End If
					End If



					If xvale = "S" Then

						xnohay2 = "N"
						cuenta=cuenta +1
						INTabla = INTabla + 1


						xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

						SESSION("ASP") = "novedmenu.asp?menu=4"
						SESSION("sEstado") = "PARA APROBAR"
						SESSION("sEstadon") = "2"
						SESSION("Form2") = "apr"
						If TIPOCPT="PD" Then
							aspeditar = "Editar.asp?tipocpt=PD&IDPEDIDO="&rsx(4)&"&ACC=PF&DOCMODAL=WKPED&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL
						ElseIf TIPOCPT="SN" Then
							aspeditar = "Editar.asp?tipocpt=SN&IDPEDIDO="&rsx(4)&"&ACC=SNN&DOCMODAL=WKPED&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL
						Else

							aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL
						End If


						''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

						aspeditar = CStr(aspeditar)
						sHref = "&nbsp;"
						'        response.write sActcon
						sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

						wNomL = "" & Capitalize(rsx(7)) & "&nbsp;-&nbsp;Paso&nbsp;" & Capitalize(rsx(16)) & "&nbsp;"
						wNomL = "" & rsx(7) & "&nbsp;-&nbsp;Paso&nbsp;" & rsx(16) & "&nbsp;"

						sImagen = wNomL


						sImagen = sImagen & "-&nbsp;" &  EsqnomPx(rsx(3), rsx(4), ofv)

						If xarc = "NO" Then
							xarc = "&nbsp;"
						End If

						sLink = sImagen
						'XTSLNKM = OFV.GenLink("", "cmb title='Autorizar documento ...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
						XTSLNKM = ofv.GenLink("", "nt title='Autorizar documento ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

						xfromusr = rsx(6)


						xxxx = getInfoUser(rsx(6))
						xfromusr = SESSION("SGU_nombre")

						ANTabla(INTabla, 0) = "CKT" & INTabla
						ANTabla(INTabla, 1) = 5
						ANTabla(INTabla, 2) = 0
						ANTabla(INTabla, 3) = dFechai
						ANTabla(INTabla, 4) = rsx(10)
						ANTabla(INTabla, 5) = "S"
						ANTabla(INTabla, 6) = "N"
						ANTabla(INTabla, 7) = xfromusr
						ANTabla(INTabla, 8) = sLink
						ANTabla(INTabla, 9) = xarc
						ANTabla(INTabla, 10) = XTSLNKM
						ANTabla(INTabla, 11) = xmenu


						ANTabla(INTabla, 12) = 2
						ANTabla(INTabla, 13) = EsqnomPx(rsx(3), rsx(4), ofv)
						ANTabla(INTabla, 14) = sLink
						ANTabla(INTabla, 15) = "0"
						ANTabla(INTabla, 16) = ""

						ANTabla(INTabla, 17) = ""
						ANTabla(INTabla, 18) = ""
						ANTabla(INTabla, 19) = ""
						ANTabla(INTabla, 20) = ""

						ANTabla(INTabla, 21) = "0"


						If cuenta >= atit(4,3) Then ' CANT REG DE LA NOVEDADES

							WSPFRMN = "seccion" & atit(4,0)
							sxlinksig =  "<A id=lnksig class=nt style='display:Xnone;'  href='javascript:void(0)' ONCLICK=document." &WSPFRMN & ".submit(); "
							sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"

							sxlinkant =  "<A id=lnkant class=nt style='display:none;' HREF=BSQcMD.asp"

							sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
							botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

							WSPFRMN = "seccion" & atit(4,0)
							'WSPFRMA = "consdocumindex.asp"
							WSPFRMA = "BSQMENU4.asp?canTregmover=" & cuenta & "&tope=" & atit(4,3) &  "&MAINAPL=" & xMAINAPL

							WSPFORMEsecc = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST'></FORM>"


							INTabla=INTabla+1
							ANTabla(INTabla, 0) = "0"
							ANTabla(INTabla, 1) = 13
							ANTabla(INTabla, 2) = 1
							ANTabla(INTabla, 3) = dFechai
							ANTabla(INTabla, 4) = 0
							ANTabla(INTabla, 5) = "S"
							ANTabla(INTabla, 6) = "N"
							ANTabla(INTabla, 7) = "MASREG"
							ANTabla(INTabla, 8) = ""
							ANTabla(INTabla, 9) = xarc
							ANTabla(INTabla, 10) = ""'"<font size=2 color=red face=tahoma>Ver Más</font>" & botones & WSPFORMEsecc
							ANTabla(INTabla, 11) = xmenu
							ANTabla(INTabla, 12) = "1"
							ANTabla(INTabla, 13) = ""
							ANTabla(INTabla, 14) = ""
							ANTabla(INTabla, 15) = ""
							ANTabla(INTabla, 16) = ""
							ANTabla(INTabla, 17) = ""
							ANTabla(INTabla, 18) = ""
							ANTabla(INTabla, 19) = ""
							ANTabla(INTabla, 20) = ""

							ANTabla(INTabla, 21) = "0"
							Exit Do
						End If

					End If



					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)
			Call ofv.cerrarconsulta(rsxp)
			Call ofv.cerrarconn(cn1)
			Call ofv.cerrarconn(cn4)
			Case 5,11
			'''    session("ModMain") = "CL"

			cuenta = 0

			''''''''''''''''''''''''''''''''
			XWHERE = " and B1.USUARIO = '" & session("usuario") & "' "
			XWHERE = XWHERE & " AND B1.CODOBJETO > 0 "

			If wkaplic = "OR" Or wkaplic = "PL" Then

				XWHERE = XWHERE & " AND B1.TIPOBJETO = 'RSK' "

			Else

				XWHERE = XWHERE & " AND B1.TIPOBJETO <> 'RSK' "

			End If

			XORDER = " group by b1.codobjeto, d.cliente ORDER BY fecha desc "


			StrSql = ""
			xcantreg=atit(5,3)
			StrSql = StrSql & "SELECT top " & xcantreg & " b1.CODOBJETO, max(b1.TIPOBJETO), max(finicio) as fecha  "
			StrSql = StrSql & " FROM OBJALOGDET B1"
			StrSql = StrSql & " inner join  objprty b on (b1.CODOBJETO=b.codobjeto and b.estado <> 4) "
			StrSql = StrSql & " inner join objetos c on (b1.CODOBJETO=c.codobjeto and b1.tipobjeto=c.tipobjeto) "

			if sbase="K" then
			  StrSql = StrSql & " inner join  TIPOBJETO Tb on (b.TIPOBJETO=Tb.TIPOBJETO and tb.tipocpt ='K') "
			else
			 StrSql = StrSql & " inner join  TIPOBJETO Tb on (b.TIPOBJETO=Tb.TIPOBJETO and tb.tipocpt <>'K') "
			end if

			StrSql = StrSql & " inner join clienteplan d on (c.objun=d.id)  "

			if  session("UN")=1 then
			 StrSql = StrSql & " WHERE  TACC = 'CONSEN'  " & XWHERE
			else
			 StrSql = StrSql & " WHERE c.objun=" & session("UN") & " and TACC = 'CONSEN'  " & XWHERE
			end if

			'StrSql = StrSql & " WHERE TACC = 'CONSEN'  " & XWHERE

			' XORDER = " group by b1.codobjeto ORDER BY fecha desc "

			StrSql = StrSql & XORDER

			' RESPONSE.WRITE strsql & "<BR>"

			Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If rsx.EOF Then
				xnohay1 = "S"

			Else

				xvinoant = "no"
				Do Until rsx.EOF
					'response.write rsx(0) & "<br>"

					xvale = "N"



					If wkaplic = "OR" Or wkaplic = "PL" Then


						StrSql = "SELECT * FROM ORRSKCIRCUITO "
						StrSql = StrSql & " Where OBJID = 0" & rsx(0)
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

						WKSPTIPOOBJ = "RSK"


					Else

						StrSql = "SELECT * FROM objetos a "
						' StrSql =  strsql & "inner join  objprty b on (a.codobjeto=b.codobjeto and b.estado <> 4) "
						StrSql = StrSql & " Where a.codobjeto = 0" & rsx(0)
						'if clng(rsx(0))=2655 then
						'  rESPONSE.WRITE StrSql & "<BR>"
						'                           end if
                      '  if xmenu=5 then    RESPONSE.WRITE STRSQL & "<BR>"
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If



						WKSPTIPOOBJ = "VER"

					End If

					If session("UN") <> 1 Then
						If rsxp(7) & "" <> session("UN") & "" Then
							'response.write "ENTRO" & "<br>"
							xvale = "N"
						End If
					End If

					Call ofv.cerrarconsulta(rsxp)



					If xvale = "S" Then


						cuenta = cuenta + 1

						xnohay1 = "N"
						INTabla = INTabla + 1


						xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


						sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
						sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
						sHref = "&nbsp;"
						' response.write sActcon
						''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

						sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


						sTipobjeto = EsqnomPx(WKSPTIPOOBJ, rsx(0), ofv)






						wmsgx = sTipobjeto & ""
						wNomL = "&nbsp;&nbsp;" & wmsgx & "&nbsp;"



						sImagen = wNomL

						If xarc = "NO" Then
							xarc = "&nbsp;"
						End If

						sLink = sImagen

						XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

						''' XTSLNKM = wNomL

						xfromusr = "COREMANAGER"


						xxxx = getInfoUser(xfromusr)

						xfromusr = SESSION("SGU_nombre")


						ANTabla(INTabla, 0) = "0"
						ANTabla(INTabla, 1) = 3
						ANTabla(INTabla, 2) = 1
						ANTabla(INTabla, 3) = dFechai
						ANTabla(INTabla, 4) = 0
						ANTabla(INTabla, 5) = "S"
						ANTabla(INTabla, 6) = "N"
						ANTabla(INTabla, 7) = xfromusr
						ANTabla(INTabla, 8) = sLink
						ANTabla(INTabla, 9) = xarc
						ANTabla(INTabla, 10) = XTSLNKM
						ANTabla(INTabla, 11) = xmenu
						ANTabla(INTabla, 12) = "1"
						ANTabla(INTabla, 13) = ""
						ANTabla(INTabla, 14) = ""
						ANTabla(INTabla, 15) = ""
						ANTabla(INTabla, 16) = ""
						ANTabla(INTabla, 17) = ""
						ANTabla(INTabla, 18) = ""
						ANTabla(INTabla, 19) = ""
						ANTabla(INTabla, 20) = ""

						ANTabla(INTabla, 21) = "0"

					End If

					If cuenta >= atit(xmenu,3)   Then  ' CANT REG DE LA NOVEDADES

						WSPFRMN = "seccion" & atit(5,0)
						sxlinksig =  "<A id=lnksig class=nt style='display:Xnone;'  href='javascript:void(0)' ONCLICK=document." &WSPFRMN & ".submit(); "
						sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"

						sxlinkant =  "<A id=lnkant class=nt style='display:none;' HREF=BSQcMD.asp"

						sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
						botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

						WSPFRMN = "seccion" & atit(5,0)
						'WSPFRMA = "consdocumindex.asp"
						WSPFRMA = "BSQUltCons.asp?canTregmover=" & cuenta & "&tope=" & atit(5,3) &  "&MAINAPL=" & xMAINAPL

						WSPFORMEsecc = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST'></FORM>"


						INTabla=INTabla+1
						ANTabla(INTabla, 0) = "0"
						ANTabla(INTabla, 1) = 13
						ANTabla(INTabla, 2) = 1
						ANTabla(INTabla, 3) = dFechai
						ANTabla(INTabla, 4) = 0
						ANTabla(INTabla, 5) = "S"
						ANTabla(INTabla, 6) = "N"
						ANTabla(INTabla, 7) = "MASREG"
						ANTabla(INTabla, 8) = ""
						ANTabla(INTabla, 9) = xarc
						ANTabla(INTabla, 10) = ""'"<font size=2 color=red face=tahoma>Ver Más</font>" & botones & WSPFORMEsecc
						ANTabla(INTabla, 11) = xmenu
						ANTabla(INTabla, 12) = "1"
						ANTabla(INTabla, 13) = ""
						ANTabla(INTabla, 14) = ""
						ANTabla(INTabla, 15) = ""
						ANTabla(INTabla, 16) = ""
						ANTabla(INTabla, 17) = ""
						ANTabla(INTabla, 18) = ""
						ANTabla(INTabla, 19) = ""
						ANTabla(INTabla, 20) = ""

						ANTabla(INTabla, 21) = "0"
						Exit Do
					End If


					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)



			Case 6

			'''    session("ModMain") = "WK"
			SESSION("Segce") = "S"
			SESSION("Segco") = "S"

			' RESPONSE.WRITE "6"

			Set cn1 = ofv.conectar(ofv.strconn5)
			Set cn4 = ofv.conectar(ofv.strconn6)

			xtipos = " OBJTIPO in ('RSK'"
			separ = ","
			StrSql = "SELECT tipobjeto FROM tipobjeto "
			Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If Not rs.EOF Then
				Do While Not rs.EOF
					xtipos = xtipos & separ & "'" & rs(0) & "'"
					separ = ","
					rs.movenext
				Loop
			Else
				xtipos = xtipos & "' '"
			End If
			xtipos = xtipos & ")"

			SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & SESSION("Usuario") & "') OR ("

			xxxx = getInfoUser(SESSION("Usuario"))
			xusrlog = SESSION("SGU_nombre")




			If xusrlog <> "NO" Then


				SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & SESSION("Usuario") & "') ) "
				SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

				'''''''''''''''''''''''''''''''''''''''''''
				If SESSION("DELLOG") = "S" Then

					XMFTGRP = Split(SESSION("SegGrpLst"), Chr(9))
					xcount = UBound(XMFTGRP) + 1
					XMFSEP = ""
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

					For XMFI = 0 To UBound(XMFTGRP)
						SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
						XMFSEP = ","
					Next

					SEGWHERE = SEGWHERE & ") ) )"

				Else
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
				End If


				''''''''''''''''''''''''''''''''''''''''''

				'''   IF  RS(21) > 0 THEN


				If CStr(SESSION("SGU_idcostcenter") & "") <> "0" Then
					SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
					SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & SESSION("SGU_idcostcenter") & "' OR SUBCODIGO = '" & SESSION("SGU_idcostcenter") & "' ) "
					SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
					SEGWHERE = SEGWHERE & " ) ) "
				End If


			End If

			SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

			StrSql = "SELECT * FROM CKTVIEW " & SEGWHERE

		' RESPONSE.WRITE STRSQL & "4444<BR>"
			Set rsx = ofv.crearconsultaEx(StrSql, cn1, 1, parametros)
			If rsx.EOF Then
				xnohay2 = "S"
			Else

				Do Until rsx.EOF

					xvale = "N"

					'   Strsql  = "SELECT * FROM objetos "
					'   Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
					'   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					'   if  not rsxp.eof then
					'       xvale = "S"
					'   end if
					'   call ofv.cerrarconsulta(rsxp)

					If rsx(3) = "RSK" Then

						StrSql = "SELECT * FROM ORRSKCIRCUITO "
						StrSql = StrSql & " Where OBJID = 0" & rsx(4)
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)


					Else



						StrSql = "SELECT * FROM objetos "
						StrSql = StrSql & " Where codobjeto = 0" & rsx(4)

						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

					End If


					If xvale = "S" Then

						xnohay2 = "N"

						INTabla = INTabla + 1

						StrSql = "SELECT archivo FROM tipobjeto "
						StrSql = StrSql & " Where tipobjeto = '" & rsx(3) & "'"

						Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rs.EOF Then
							'         xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
							'      else
							'          xarc = "NO"
						End If
						Call ofv.cerrarconsulta(rs)

						xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

						SESSION("ASP") = "novedmenu.asp?menu=4"
						SESSION("sEstado") = "PARA APROBAR"
						SESSION("sEstadon") = "2"
						SESSION("Form2") = "apr"

						aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8)


						''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

						aspeditar = CStr(aspeditar)
						sHref = "&nbsp;"
						'        response.write sActcon
						sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

						''''''wNomL = "" & Capitalize(rsx(7)) & "&nbsp;-&nbsp;PASO&nbsp;" & Capitalize(rsx(16)) & "&nbsp;"
						wNomL = "" & rsx(7) & "&nbsp;-&nbsp;PASO&nbsp;" & rsx(16) & "&nbsp;"

						sImagen = wNomL


						sImagen = sImagen & "-&nbsp;" & EsqnomPx(rsx(3), rsx(4), ofv)


						If xarc = "NO" Then
							xarc = "&nbsp;"
						End If

						sLink = sImagen
						XTSLNKM = ofv.GenLink("", "istt title='Ver mensaje ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")


						xfromusr = rsx(6)


						xxxx = getInfoUser(rsx(6))
						xfromusr = SESSION("SGU_nombre")


						'' ANTabla(INTabla,0) = "0"
						ANTabla(INTabla, 0) = "MISPEN" & INTabla
						ANTabla(INTabla, 1) = 5
						ANTabla(INTabla, 2) = 0
						ANTabla(INTabla, 3) = dFechai
						ANTabla(INTabla, 4) = rsx(10)
						ANTabla(INTabla, 5) = "S"
						ANTabla(INTabla, 6) = "N"
						ANTabla(INTabla, 7) = xfromusr
						ANTabla(INTabla, 8) = sLink
						ANTabla(INTabla, 9) = xarc
						ANTabla(INTabla, 10) = XTSLNKM
						ANTabla(INTabla, 11) = xmenu
						ANTabla(INTabla, 12) = 2
						ANTabla(INTabla, 13) = EsqnomPx(rsx(3), rsx(4), ofv)
						ANTabla(INTabla, 14) = sLink
						ANTabla(INTabla, 15) = "0"
						ANTabla(INTabla, 16) = ""

						ANTabla(INTabla, 17) = ""
						ANTabla(INTabla, 18) = ""
						ANTabla(INTabla, 19) = ""
						ANTabla(INTabla, 20) = ""

						ANTabla(INTabla, 21) = "0"

					End If



					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)
			Call ofv.cerrarconn(cn1)
			Call ofv.cerrarconn(cn4)
			''''''''''''''''''''''''''''''''''''''''''

			Case 7

			'''    session("ModMain") = "CL"

			cuenta = 0


			If session("UN") = 1 Then
				StrSql = ""
				StrSql = StrSql & "SELECT a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente "
				StrSql = StrSql & " from objetos a "
				StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
				StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
				StrSql = StrSql & " where b.fecvto > 0 order by b.fecvto "
			Else
				StrSql = ""
				StrSql = StrSql & "SELECT a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente "
				StrSql = StrSql & " from objetos a "
				StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
				StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
				StrSql = StrSql & " where b.fecvto > 0 and a.objun=" & session("UN") & " order by b.fecvto "
			End If

			'RESPONSE.WRITE STRSQL & "<BR>"

			Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If rsx.EOF Then
				xnohay1 = "S"

			Else

				xvinoant = "no"

				Do Until rsx.EOF


					xvale = "N"



					If wkaplic = "OR" Or wkaplic = "PL" Then




					Else


						xvale = "S"
						xMAINAPL = rsx("APLIC")



						WKSPTIPOOBJ = "VER"

					End If

					If session("UN") <> 1 Then
						If rsx(4) & "" <> session("UN") & "" Then xvale = "N"
					End If

					'' agregado 6-11-2014
					dfecvto = rsx(2)
					dfecvto = CLng(dfecvto)
					dfechadia = CLng(Date) - CLng(CDate("01-01-1900"))
					'' RESPONSE.WRITE dFechadia & " dFechadia " & "<BR>"
					'' fin agregado 6-11-2014

					If xvale = "S" Then

						'' agregado 6-11-2014

						'' RESPONSE.WRITE dfecvto & " dfecvto "  & "<BR>"
						'' RESPONSE.WRITE dFechadia & " dFechadia " & "<BR>"

						If dfecvto > dfechadia Then

							StrSql = ""
							StrSql = StrSql & "SELECT tipobjeto, avisoant from tipobjeto where aplic in('EM','GLB') "
							StrSql = StrSql & " and tipobjeto = '" & rsx(1) & "' "
							''RESPONSE.WRITE STRSQL & "<BR>"
							Set rsy = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
							If Not rsy.EOF Then
								'' RESPONSE.WRITE " linea 2190 "  & "<BR>"
								davisoant = 0
								davisoant = rsy(1)
								dfecaviso = rsx(2) - davisoant

								''RESPONSE.WRITE davisoant & " davisoant "  & "<BR>"
								''RESPONSE.WRITE dfecvto & " dfecvto "  & "<BR>"

								dfecaviso = CLng(dfecaviso)

								''RESPONSE.WRITE dfecaviso & " dfecaviso "  & "<BR>"
								''RESPONSE.WRITE dFechadia & " dFechadia " & "<BR>"

								If (dfecaviso <= dfechadia And dfecvto >= dfechadia) Then
									'''''''''''''''''''''''''''''''If dfecaviso > dFechadia Then
									diferavis = 0
									diferavis = dfecaviso - dfechadia
									xdife = 0
									xdife = diferavis - davisoant

									''RESPONSE.WRITE diferavis & " diferavis "  & "<BR>"
									''RESPONSE.WRITE xdife & " xdife "  & "<BR>"

									'' if diferavis = davisoant then




									'' fin agregado 6-11-2014

									cuenta = cuenta + 1

									xnohay1 = "N"
									INTabla = INTabla + 1


									xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


									sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
									sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
									sHref = "&nbsp;"
									' response.write sActcon
									''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

									sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


									sTipobjeto = EsqnomPx(WKSPTIPOOBJ, rsx(0), ofv)

									xtfechavig = CLng("0" & rsx(2))
									xtfechavig = GTFechadorsh(xtfechavig)



									wmsgx = sTipobjeto & ""
									wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;&nbsp;&nbsp;<font color='DARKRED' >(" & xtfechavig & ")</FONT>"



									sImagen = wNomL

									If xarc = "NO" Then
										xarc = "&nbsp;"
									End If

									sLink = sImagen

									XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

									''' XTSLNKM = wNomL

									xfromusr = "COREMANAGER"


									xxxx = getInfoUser(xfromusr)

									xfromusr = SESSION("SGU_nombre")


									ANTabla(INTabla, 0) = "0"
									ANTabla(INTabla, 1) = 3
									ANTabla(INTabla, 2) = 1
									ANTabla(INTabla, 3) = dFechai
									ANTabla(INTabla, 4) = 0
									ANTabla(INTabla, 5) = "S"
									ANTabla(INTabla, 6) = "N"
									ANTabla(INTabla, 7) = xfromusr
									ANTabla(INTabla, 8) = sLink
									ANTabla(INTabla, 9) = xarc
									ANTabla(INTabla, 10) = XTSLNKM
									ANTabla(INTabla, 11) = xmenu
									ANTabla(INTabla, 12) = "1"
									ANTabla(INTabla, 13) = ""
									ANTabla(INTabla, 14) = ""
									ANTabla(INTabla, 15) = ""
									ANTabla(INTabla, 16) = ""
									ANTabla(INTabla, 17) = ""
									ANTabla(INTabla, 18) = ""
									ANTabla(INTabla, 19) = ""
									ANTabla(INTabla, 20) = ""

									ANTabla(INTabla, 21) = "0"
									'' agregado 6-11-2014
									''End If

								End If

							End If

						End If
						'' fin agregado 6-11-2014



					End If

					If cuenta >= atit(7,3) Then Exit Do

					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)


			'''''''''''''''''''''''''''''''''''''''''


			''''''''''''''''''''''''''''''''''''''''''

			Case 8

			'''    session("ModMain") = "CL"


			StrSql = "SELECT distinct a.codobjeto,a.tipobjeto FROM OBJPRTY a "
			StrSql = StrSql & " INNER JOIN  OBJOBSCMT B ON (b.codsol=9999 and B.tipo='CMD' AND B.ACTPEND is null  AND B.CODOBJETO=a.CODOBJETO AND B.TIPOBJETO=a.TIPOBJETO ) "
			StrSql = StrSql & " INNER JOIN  TIPOBJETO C ON (C.COMENTARIOS='S'  AND A.TIPOBJETO=C.TIPOBJETO ) "
			'StrSql = StrSql & " INNER JOIN  objsdkdet D ON (D.tipoSEG='9' AND D.CODOBJETO=0 AND A.TIPOBJETO=D.TIPOBJETO ) "
			StrSql = StrSql & " WHERE  (a.ESTADO = 3 )"
			StrSql = StrSql & " And (a.autor='"& SESSION("usuario") &"')"
			'StrSql = StrSql & " And ("
			'StrSql = StrSql & " (D.GRPUSR='USR-" & SESSION("usuario") & "')"
			'StrSql = StrSql & " or "
			'StrSql = StrSql & " (D.GRPUSR IN ('GRP-" & Session("SegGrp") &"'"
			'wktgrupos = SPLIT(session("SegGrpLst"), CHR(9))

			'FOR wkti = 0 TO UBOUND(wktgrupos)


			'StrSql = StrSql & ",'GRP-" &  wktgrupos(wkti)  &"'"


			'NEXT



			' StrSql = StrSql & ")"
			' RESPONSE.WRITE STRSQL

			cuenta = 0



			Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If rsx.EOF Then
				xnohay1 = "S"

			Else

				xvinoant = "no"
				Do Until rsx.EOF


					xvale = "N"



					If wkaplic = "OR" Or wkaplic = "PL" Then


						StrSql = "SELECT * FROM ORRSKCIRCUITO "
						StrSql = StrSql & " Where OBJID = 0" & rsx(0)
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

						WKSPTIPOOBJ = "RSK"


					Else

						StrSql = "SELECT * FROM objetos "
						StrSql = StrSql & " Where codobjeto = 0" & rsx(0)
						''    RESPONSE.WRITE STRSQL & "<BR>"
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						If Not rsxp.EOF Then
							xvale = "S"
							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)


						WKSPTIPOOBJ = "VER"

					End If




					If xvale = "S" Then


						cuenta = cuenta + 1

						xnohay1 = "N"
						INTabla = INTabla + 1


						xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"

						'INC 1373
						'Se agrega PEND=S para que muestre botón Confirmar en comentarios al documento
						'al autor del documento y a los autorizantes

						sActcon = "ConsultarcMD.asp?PEND=S&Codigo=" & rsx(0) & "&Tipo=" & rsx(1)
						sHref = "&nbsp;"


						sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


						sTipobjeto = EsqnomPx(WKSPTIPOOBJ, rsx(0), ofv)






						wmsgx = sTipobjeto & ""
						wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



						sImagen = wNomL

						If xarc = "NO" Then
							xarc = "&nbsp;"
						End If

						sLink = sImagen

						XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

						''' XTSLNKM = wNomL

						xfromusr = "COREMANAGER"


						xxxx = getInfoUser(xfromusr)

						xfromusr = SESSION("SGU_nombre")


						ANTabla(INTabla, 0) = "0"
						ANTabla(INTabla, 1) = 13
						ANTabla(INTabla, 2) = 1
						ANTabla(INTabla, 3) = dFechai
						ANTabla(INTabla, 4) = 0
						ANTabla(INTabla, 5) = "S"
						ANTabla(INTabla, 6) = "N"
						ANTabla(INTabla, 7) = xfromusr
						ANTabla(INTabla, 8) = sLink
						ANTabla(INTabla, 9) = xarc
						ANTabla(INTabla, 10) = XTSLNKM
						ANTabla(INTabla, 11) = xmenu
						ANTabla(INTabla, 12) = "1"
						ANTabla(INTabla, 13) = ""
						ANTabla(INTabla, 14) = ""
						ANTabla(INTabla, 15) = ""
						ANTabla(INTabla, 16) = ""
						ANTabla(INTabla, 17) = ""
						ANTabla(INTabla, 18) = ""
						ANTabla(INTabla, 19) = ""
						ANTabla(INTabla, 20) = ""

						ANTabla(INTabla, 21) = "0"

					End If

					If cuenta >= atit(8,3) Then ' CANT REG DE LA NOVEDADES

						WSPFRMN = "seccion" & atit(8,0)
						sxlinksig =  "<A id=lnksig class=nt style='display:Xnone;'  href='javascript:void(0)' ONCLICK=document." &WSPFRMN & ".submit(); "
						sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"

						sxlinkant =  "<A id=lnkant class=nt style='display:none;' HREF=BSQcMD.asp"

						sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
						botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

						WSPFRMN = "seccion" & atit(8,0)
						'WSPFRMA = "consdocumindex.asp"
						WSPFRMA = "BSQcMD.asp?canTregmover=" & cuenta & "&tope=" & atit(8,3)

						WSPFORMEsecc = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST'></FORM>"


						INTabla=INTabla+1
						ANTabla(INTabla, 0) = "0"
						ANTabla(INTabla, 1) = 13
						ANTabla(INTabla, 2) = 1
						ANTabla(INTabla, 3) = dFechai
						ANTabla(INTabla, 4) = 0
						ANTabla(INTabla, 5) = "S"
						ANTabla(INTabla, 6) = "N"
						ANTabla(INTabla, 7) = "MASREG"
						ANTabla(INTabla, 8) = ""
						ANTabla(INTabla, 9) = xarc
						ANTabla(INTabla, 10) = "" '"<font size=2 color=red face=tahoma>Ver Más</font>" & botones & WSPFORMEsecc
						ANTabla(INTabla, 11) = xmenu
						ANTabla(INTabla, 12) = "1"
						ANTabla(INTabla, 13) = ""
						ANTabla(INTabla, 14) = ""
						ANTabla(INTabla, 15) = ""
						ANTabla(INTabla, 16) = ""
						ANTabla(INTabla, 17) = ""
						ANTabla(INTabla, 18) = ""
						ANTabla(INTabla, 19) = ""
						ANTabla(INTabla, 20) = ""

						ANTabla(INTabla, 21) = "0"
						Exit Do
					End If

					rsx.movenext
				Loop
			End If
			Call ofv.cerrarconsulta(rsx)


			'''''''''''''''''''''''''''''''''''''''''

			Case 10

			'''   session("ModMain") = "CL"

			xnohay1 = "N"


			sLink = "<INPUT TYPE=BUTTON VALUE='HIDE MENU' onclick='hdemenu();' ><INPUT TYPE=BUTTON VALUE='SHOW MENU' onclick='mostrarmenu();' >"


			INTabla = INTabla + 1


			ANTabla(INTabla, 0) = "0"
			ANTabla(INTabla, 1) = 3
			ANTabla(INTabla, 2) = 1
			ANTabla(INTabla, 3) = dFechai
			ANTabla(INTabla, 4) = 0
			ANTabla(INTabla, 5) = "S"
			ANTabla(INTabla, 6) = "N"
			ANTabla(INTabla, 7) = "COREMANAGER"
			ANTabla(INTabla, 8) = sLink
			ANTabla(INTabla, 9) = "&nbsp;"
			ANTabla(INTabla, 10) = "&nbsp;"
			ANTabla(INTabla, 11) = xmenu
			ANTabla(INTabla, 12) = "1"
			ANTabla(INTabla, 13) = ""
			ANTabla(INTabla, 14) = ""
			ANTabla(INTabla, 15) = ""
			ANTabla(INTabla, 16) = ""


			ANTabla(INTabla, 17) = ""
			ANTabla(INTabla, 18) = ""
			ANTabla(INTabla, 19) = ""
			ANTabla(INTabla, 20) = ""
			ANTabla(INTabla, 21) = "0"








			Case Else


		End Select

		If menuSel <> "" Then
			saleSel = "S"
		Else
			If INTabla > 0 Then
				saleSel = "S"
			Else
				menuSel = 2
				Menu = 2
			End If
		End If



	Loop
noveselec=INTabla
End Function

Function noveformat(xmenu, XMENSEL)


	XTablaN = ""

	acctitblks = 0
	acctitlin = 0


	'xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "

	xsclase = " CLASS=ietx "
	xsclase2 = "class=istt "
	'''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "

	XTablaN = XTablaN & "<div id='DTNOV" & xmenu & "' style='width:100%; "


	'If xmenu = 2 Then

	'    XTABLAN = XTablaN & " width:40%;"

	'Else

	'   XTABLAN = XTablaN & " width:30%;"

	'End If

	'********* ENCABEZADO DE SECCIONES
	XTablaN = XTablaN & "border-style:solid;border-width:3;border-color:#ffffff;' >"
	XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1     >"
	XTablaN = XTablaN & "<tr HEIGHT=25  >"
	''''XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(Capitalize(atit(xmenu)) & "", " ", "&nbsp;") &  "</TD>"
	XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(atit(xmenu,1) & "", " ", "&nbsp;") & "</TD>"
	XTablaN = XTablaN & "</TR>"
	XTablaN = XTablaN & "</table>"

	XTablaN = XTablaN & "<div id='DTNOVINT" & xmenu & "' style='"

	XTablaN = XTablaN & "border-style:solid;border-width:1;border-color:#ffffff;' >"

	XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=0  >"


	If TNTabla = 0 Then


		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(Menu,1) & "&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"

		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"


	Else
		XTablaNF=""
		For INTabla = 1 To TNTabla

			'    acctitlin  = acctitlin + 2


			xtmcod = ANTabla(INTabla, 0)
			xtipom = ANTabla(INTabla, 1)
			xtpri = ANTabla(INTabla, 2)
			xtfvig = ANTabla(INTabla, 3)
			xthvig = ANTabla(INTabla, 4)
			xtsadj = ANTabla(INTabla, 5)
			xtslnk = ANTabla(INTabla, 6)
			xfromusr = ANTabla(INTabla, 7)
			xtmsg = ANTabla(INTabla, 8)
			xtmsgimg = ANTabla(INTabla, 9)
			xtmsglnk = ANTabla(INTabla, 10)
			xtmsgTYPE = ANTabla(INTabla, 11)



			xfclas = "dt"
			sImage = "&nbsp;"

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

			sglmod = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
			sImage = ofv.GenLink("", xfclas & " title='" & sgltit & "' ", "", "", "", "", sglmod, "si")

			If xtfvig = dFechai Then
				XMARQ = sglmar
			Else
				XMARQ = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
			End If

			If xtipom = 5 Then
				xtfechavig = xthvig
				XTCOUNT = InStr(xtfechavig, " ")
				If XTCOUNT > 0 Then
					xtfechavig = Left(xtfechavig, XTCOUNT)
					xtfechavig = GTFechadorsh(CLng(CDate(xtfechavig)) - CLng(CDate("01-01-1900")))
				End If
			Else
				xtfechavig = GTFechadorsh(xtfvig)
				XTCOUNT = InStr(xtfechavig, " ")
				If XTCOUNT > 0 Then xtfechavig = Left(xtfechavig, XTCOUNT)


			End If

			Xnlnk = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
			Xnadj = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

			If xtsadj = "S" Then Xnadj = sgladj
			If xtslnk = "S" Then Xnlnk = sgllnk









			xtmsgVRS = "0" & ANTabla(INTabla, 12)
			xtmsgVRS = CInt(xtmsgVRS)
			xtmsgTIT = ANTabla(INTabla, 13)
			xtmsgCOP = ANTabla(INTabla, 14)
			xtmsgimg = ANTabla(INTabla, 15)
			xtmsgVID = ANTabla(INTabla, 16)


			xtmsgVAENC = ANTabla(INTabla, 17)
			xtmsgVAFORO = ANTabla(INTabla, 18)
			xtmsgENCCOD = ANTabla(INTabla, 19)
			xtmsgFOROCOD = ANTabla(INTabla, 20)

			xtmsgdoccod = ANTabla(INTabla, 21)


			If xtmsgVRS = 2 Then

				acctitlin = acctitlin + 2

				'xtmsglnk    = replace(xtmsglnk,"XXMSGXX","<font STYLE='color:#990000;font-size:10pt;font-weight:bold;' >" & xtmsgTIT & "</font>")
				xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

				'  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
				'  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
				'  XTABLAN = XTablaN &  "</TR>"

				XTablaN = XTablaN & "<tr HEIGHT='20' >"
				XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
				For yy = 1 To 6
					XTablaN = XTablaN & "<td class=UT width='15%'  >&nbsp;</TD>"
				Next
				XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
				XTablaN = XTablaN & "</TR>"

				''  TIPO DE NOVEDAD NOTICIA DEL DIA
				If xtipom = "1" Or xtipom = "2"  Then

					xSolasp = ""
					xSolasp = xSolasp & "novedmenuview.asp?xxmsg=" & CStr(xtmcod) & "&Pantalla=MAIN"




					XTablaNF = XTablaNF & "<form name='VNGP" & CStr(xtmcod) & "' id='VNGP" & CStr(xtmcod) & "' action='" & xSolasp & "' "
					XTablaNF = XTablaNF & " method=post target='fesquemas' style='display:none;' >"
					' XTablaN = XTablaN & " method=post target='_new' style='display:none;' >"
					XTablaNF = XTablaNF & "</form>"



					xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

					XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"


					If SESSION("APLMDCLI") = "BGAL" Or SESSION("APLMDCLI") = "GP"  Then
						XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8   ><MARQUEE>" & xtmsglnk & "</MARQUE></a></TD>"

					Else

						XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8   >" & xtmsglnk & "</a></TD>"


					End If


					XTablaN = XTablaN & "</TR>"





					'' Otros tipos de novedad
				Else



					XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"
					XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8  >" & xtmsglnk & "</TD>"

					XTablaN = XTablaN & "</TR>"

				End If


				If xtipom = "9" Then

					acctitlin = acctitlin + 2

					XTMSGSEND = ""
					XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP

					XTablaN = XTablaN & "<tr  >"
					XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B>"
					XTablaN = XTablaN & "<font color='#AA0000' size=2 >&#0187;</font>"
					XTablaN = XTablaN & "<font color='#0000AA' size=1 >&#0124;</font>"
					XTablaN = XTablaN & "</B></TD>"
					XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
					XTablaN = XTablaN & XTMSGSEND & "</TD>"
					XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
					XTablaN = XTablaN & "</TR>"


					XTMSGSEND = ""
					XTMSGSEND = XTMSGSEND & "<font color=#888888 style='font-weight:bold;' size=1 >recomendado por " & ANTabla(INTabla, 7) & "</font>"

					XTablaN = XTablaN & "<tr HEIGHT='20' >"
					XTablaN = XTablaN & "<td class=UT align='right' colspan=8 STYLE='background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
					XTablaN = XTablaN & "</TR>"

				End If


				'''''''''''''''''''''''''''''
				XTablaN = XTablaN & "<tr HEIGHT='20' >"
				XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
				XTablaN = XTablaN & "</TR>"


				If xtmsgdoccod <> "0" Then



					StrSql = "Select * From objinfo where codobjeto = 0" & xtmsgdoccod & " "
					'' response.write strsql
					Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs1.EOF Then
						'aca esta el error

						''''' AGREGADO PARA MEJORAS VERSION 4.6
						xtmsgCOPP = rs1(6)
						If xtmsgCOPP <> "" Then
							xtmsgCOP = xtmsgCOP & "<br><br>" & rs1(6)
						End If
						''''' AGREGADO PARA MEJORAS VERSION 4.6 FIN

						xtmsgCOP = xtmsgCOP & "<br><br>" & rs1(4) '& xtmsgdoccod
						xtmsgimg = rs1(7)
						xtmsgVID = rs1(9)

					End If

					StrSql = "Select * From objinfoIMG  where codobjeto = 0" & xtmsgdoccod & " ORDER BY objSECUEN, OBJID"

					wspimgdat = 6

				Else

					StrSql = "Select * From NOVEDADESIMG  where NOVEDAD = '" & xtmcod & "' ORDER BY SECUEN, OBJID"
					wspimgdat = 5
				End If



				If (xtmsgimg & "") = "1" Then

					'   response.write Strsql

					Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs1.EOF Then



						XTablaN = XTablaN & "<tr bgcolor='#F0F0F0' >"
						XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

						XTablaN = XTablaN & "<table cellspacing=0 cellpadding=1  width=100% >"
						XTablaN = XTablaN & "<tr   >"


						XSFSTYLE = " STYLE='display:block;' "

						Do While Not rs1.EOF

							acctitlin = acctitlin + 10

							xCelda = ""
							xCelda = xCelda & "<div id='DIMGN" & INTabla & "' style='margin-top:1;margin-left:1;"
							xCelda = xCelda & "width:98%;overflow:auto;"
							xCelda = xCelda & ""
							xCelda = xCelda & "' >" & Chr(13)
							xCelda = xCelda & "<IMG style='max-width:98%;border-style:solid;border-width:1;border-color:#888888;' SRC=" & Chr(34) & rs1(wspimgdat) & Chr(34) & ">"
							xCelda = xCelda & "</div>" & Chr(13)

							XTablaN = XTablaN & "<td CLASS=UT align='center' " & XSFSTYLE & " width=100% >" & xCelda & "</TD>"
							XSFSTYLE = " STYLE='display:none;' "


							rs1.movenext
						Loop

						XTablaN = XTablaN & "</TR>"
						XTablaN = XTablaN & "</table>"

						XTablaN = XTablaN & "</TD>"
						XTablaN = XTablaN & "</TR>"


						XTablaN = XTablaN & "<tr HEIGHT='20' >"
						XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
						XTablaN = XTablaN & "</TR>"


					End If






				ElseIf (xtmsgimg & "") = "2" Then

					acctitlin = acctitlin + 10


					XTablaN = XTablaN & "<tr bgcolor='#F0F0F0' >"
					XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
					XTablaN = XTablaN & "<a class=istt target='_blank' href=" & xtmsgVID & "> " & xtmsgVID & "</a>"
					'  XTablaN = XTablaN & xtmsgVID
					'   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
					XTablaN = XTablaN & "</TD>"
					XTablaN = XTablaN & "</TR>"


					XTablaN = XTablaN & "<tr HEIGHT='20' >"
					XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
					XTablaN = XTablaN & "</TR>"

				End If

				acctitlin = acctitlin + 2

				'**** COPETE NOTICIA DEL DIA
				XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
				XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
				XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP
				'''''response.write xtmsgCOP & " xtmsgCOP - mensaje "
				XTablaN = XTablaN & "<tr  >"
				XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B><font color='#AA0000' size=2 >&#0149;</font></B></TD>"
				XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='text-align:justify;' >"
				' LINK VER MAS
				'     If xtipom = "2" or true Then
				' url = "novedmenuview.asp?xxmsg=" & xtmcod


				If session("APLMDCLI")="GP" Then
					XTablaN = XTablaN & XTMSGSEND & "<br> <br>"
				Else
					url="novedmenuview.asp?xxmsg=" & CStr(xtmcod) & "&Pantalla=MAIN"
					sFuncion = " onclick=" & Chr(34) & "document.getElementById('VNGP" & CStr(xtmcod) & "').submit();WKlevanta('" & CStr(xtmcod) & "',this);" & Chr(34)
					xtmsglnk = ofv.GenLink("", "istt style='cursor:pointer ;' title=' Ver Mas ...' ", "", "", "", sFuncion, "Ver Mas", "si")
					If xmenu = 4 Then
						XTablaN = XTablaN & XTMSGSEND & "<br> <br>"
					Else
						If (xtipom = "1" Or xmenu = "2" )  And tieneMensaje(xtmcod) = 1 Then
							XTablaN = XTablaN & XTMSGSEND & "<br> <br>" & xtmsglnk
						Else
							XTablaN = XTablaN & XTMSGSEND & "<br> <br>"
						End If
					End If
					XTablaNF = XTablaNF & "<form name='VNGP" & CStr(xtmcod) & "' id='VNGP" & CStr(xtmcod) & "' action='" & url & "' "
					XTablaNF = XTablaNF & " method=post target='fesquemas' style='display:none;' >"

					XTablaNF = XTablaNF & "</form>"
				End If

				'  Else
				'  XTablaN = XTablaN & XTMSGSEND & "</TD>"
				'  End If
				' XTablaN = XTablaN & "</TD>"







				''' ********
				XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
				XTablaN = XTablaN & "</TR>"




				''''' PREGUNTA SI VA ENCUESTA '''''''''''''''''''''''

				'    xtmsgVAENC      = ANTabla(INTabla,17)
				'    xtmsgVAFORO     = ANTabla(INTabla,18)
				'    xtmsgENCCOD     = ANTabla(INTabla,19)
				'    xtmsgFOROCOD    = ANTabla(INTabla,20)



				If xtmsgVAENC = "S" Then

					acctitlin = acctitlin + 2

					StrSql = "Select * From emencuestas  where objid = 0" & xtmsgENCCOD
					Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs1.EOF Then

						'''''''xtmsgTITENC = Capitalize(rs1(17) & "")
						xtmsgTITENC = rs1(17) & ""
						xtmsgENCDER = "S"


						If (rs1(2) & "") = "1" Then
							StrSql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & xtmsgENCCOD
							StrSql = StrSql & " Order By BLKSECUEN,objid DESC"
							Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
							If Not rs1.EOF Then
								sblktvoto = rs1(4) & ""
								sblkcodigo = rs1(0) & ""
								If sblktvoto = "2" Then
									xtmsgENCDER = "N"
									''''''xtmsgTITENC = Capitalize(rs1(3) & "")
									xtmsgTITENC = rs1(3) & ""



									XTablaN = XTablaN & "<tr  >"
									XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' >&nbsp;</TD>"
									XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='color:#000000;' align='center' ><br>"

									XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=1  style='background-color:#eeeeee;border-style:solid;border-width:1;border-color:#dddddd;' >"

									XTablaN = XTablaN & "<tr HEIGHT=20  >"
									XTablaN = XTablaN & "<td class=vtI COLSPAN=3  >Encuesta</TD>"
									XTablaN = XTablaN & "</TR>"


									XTablaN = XTablaN & "<tr HEIGHT=20  >"
									XTablaN = XTablaN & "<td class=utI width=85% style='font-size:8pt;font-weight:normal;' >" & xtmsgTITENC & "</TD>"

									WSPDIVN = "DENC" & xtmsgENCCOD

									StrSql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & xtmsgENCCOD
									StrSql2 = StrSql2 & " and BLKcodigo = 0" & sblkcodigo
									StrSql2 = StrSql2 & " Order By ITMSECUEN,objid DESC"

									Set rs2 = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)
									Do While (Not rs2.EOF)

										acctitlin = acctitlin + 2


										WSPFRMN = "FENC" & xtmsgENCCOD & "IT" & rs2(0)
										WSPFRMA = "EMENCUESTASACT.ASP?TP=ES&OP=ACT&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
										WSPFRMR = "EMENCUESTASACT.ASP?TP=ES&OP=RES&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
										XTablaN = XTablaN & "<td class=ut width=5% align='center' >"
										WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"
										XTablaN = XTablaN & "<input type='button' style='width:45px;background-color:#cc0000;color:#ffffff;font-weight:bold;border-style:solid;border-width:1;border-color:#888888;' "
										XTablaN = XTablaN & "  value='&nbsp;" & rs2(4) & "&nbsp;' "
										XTablaN = XTablaN & "  onclick='"
										'  XTABLAN = XTablaN &  "alert("" voto:" & rs2(4) & " "");"
										'  XTABLAN = XTablaN &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
										XTablaN = XTablaN & "document." & WSPFRMN & ".submit();"
										XTablaN = XTablaN & "document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";"
										XTablaN = XTablaN & "' >"
										XTablaN = XTablaN & "</TD>" & WSPFORME
										rs2.movenext
									Loop

									acctitlin = acctitlin + 2

									xCelda = ""
									xCelda = xCelda & "<div id='" & WSPDIVN & "' style='z-index:1;margin-top:1;margin-left:1;display:none;"
									xCelda = xCelda & "width:99%;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
									xCelda = xCelda & "background-color:#F9F9F9;"
									xCelda = xCelda & "' >" & Chr(13)
									xCelda = xCelda & "&nbsp;"
									xCelda = xCelda & "</div>" & Chr(13)


									XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"
									XTablaN = XTablaN & "</TR>"

									XTablaN = XTablaN & "<tr HEIGHT=20  >"
									XTablaN = XTablaN & "<td class=utI width=85%  >&nbsp;</TD>"
									XTablaN = XTablaN & "<td class=vt COLSPAN=2 ALIGN='CENTER' ><A CLASS=NT HREF='" & WSPFRMR & "' TARGET='fesquemas' "
									XTablaN = XTablaN & " ONCLICK='document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";' "
									XTablaN = XTablaN & "><B>Resultados</B></TD>"
									XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"

									XTablaN = XTablaN & "</TR>"

									XTablaN = XTablaN & "<tr   >"
									XTablaN = XTablaN & "<td class=vtI COLSPAN=4 >&nbsp;</TD>"
									XTablaN = XTablaN & "</TR>"



									XTablaN = XTablaN & "</table>" & xCelda
									XTablaN = XTablaN & "</TD>"
									XTablaN = XTablaN & "<td class=ut   >&nbsp;</TD>"
									XTablaN = XTablaN & "</TR>"





								End If

							End If


						End If





					End If

				End If


				''''''''''''''''''''''''''''

				XTablaN = XTablaN & "<tr HEIGHT='20' >"
				XTablaN = XTablaN & "<td class=UT colspan=8 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >&nbsp;</TD>"
				XTablaN = XTablaN & "</TR>"











			Else

				acctitlin = acctitlin + 1


				''''  xtmsg    = replace(ltrim(xtmsg & "")," ","&nbsp;")
				xtmsg = LTrim(xtmsg & "")
				xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsg & "")


				If xtmcod = "0" Then

					XTMSGSEND = "<font color=#000000 >" & xtmsglnk & "</font>"

				Else

					XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
					XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
					XTMSGSEND = XTMSGSEND & "<font color=#000000 >&nbsp;" & xtmsglnk & "</font>"

				End If



				XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " HEIGHT='25' >"
				XTablaN = XTablaN & "<td class=" & xfclas & " align='center' >" & xtmsgimg & "</TD>"
				XTablaN = XTablaN & "<td class=uti colspan=6 STYLE='text-align:justify;' >" & XTMSGSEND & "</TD>"
				XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
				XTablaN = XTablaN & "</TR>"


				If xtmcod <> "0" Then

					' acctitlin  = acctitlin + 1


					XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=5 >"
					XTablaN = XTablaN & "<td class=ut colspan=8 valign=top align=right style='border-top-style:solid;border-top-width:1;border-top-color:#eeeeee;' >&nbsp;</TD>"
					XTablaN = XTablaN & "</TR>"

				End If


			End If


			If xTrColor = "#FFFFFF" Then
				xTrColor = "#eeeeff"
			Else
				xTrColor = "#FFFFFF"
			End If


			'  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
			'  XTABLAN = XTablaN &  "<td class=UT colspan=8  ><HR  size=1 style='width:400;color:#ff0000;' >&nbsp;</TD>"
			'  XTABLAN = XTablaN &  "</TR>"




		Next


	End If

	XTablaN = XTablaN & "<tr HEIGHT='20' >"
	XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
	XTablaN = XTablaN & "</TR>"



	XTablaN = XTablaN & "</table>"
	XTablaN = XTablaN & "<HR  size=1 style='width:350;color:#eeeeee;' >"
	XTablaN = XTablaN & "</DIV>"

	XTablaN = XTablaN & "</DIV>"  & XTablaNF

	'response.write xmenu

	If TNTabla = 0 Then
		noveformat = ""
	Else

		If xmenu = 4 Then atmensajes = TNTabla

		ttitsecc = ttitsecc + 1

		'''''atitsecciones(ttitsecc, 0) = Capitalize(atit(xmenu))
		atitsecciones(ttitsecc, 0) = atit(xmenu,1)
		atitsecciones(ttitsecc, 1) = TNTabla
		atitsecciones(ttitsecc, 2) = acctitlin + 2
		atitsecciones(ttitsecc, 3) = XTablaN
		atitsecciones(ttitsecc, 4) = 0
		atitsecciones(ttitsecc, 5) = 0

		noveformat = XTablaN
	End If


End Function


Function Favoritos(tit)
	Dim XTablaN
	StrSql = " Select *,B.ESTADO   From objmsg a "
	StrSql = StrSql & " LEFT JOIN OBJPRTY B ON(A.MSGOBJCOD= B.CODOBJETO ) "
	StrSql = StrSql & " where a.MSGTIPOMESS=1  and b.estado=3 AND usr = '" & SESSION("USUARIO") & "'"
	'response.write strsql
	xTrColor = ""
	Set rsfav = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


	    xsclase = " CLASS=ietx "
		xsclase2 = "class=istt "
		'''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "

		XTablaN ="<div id='DTNOVF' style='width:100%;xbackground-color:#dddddd;' "
        XTablaN = XTablaN & "xborder-style:solid;border-width:3;border-color:#ffffff;' >"




    if tit= "S" then
      '********* ENCABEZADO DE SECCIONES

	  		XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1     >"
	  		XTablaN = XTablaN & "<tr HEIGHT=25  >"
	  		''''XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(Capitalize(atit(xmenu)) & "", " ", "&nbsp;") &  "</TD>"
	  		XTablaN = XTablaN & "<td " & xsclase & "   >FAVORITOS</TD>"
	  		XTablaN = XTablaN & "</TR>"
	  	    XTablaN = XTablaN & "</table>"

      XTablaN = XTablaN & "<table  width=98% cellspacing=0 cellpadding=0  >"

    else
    	XTablaN = "<table  width=98% cellspacing=0 cellpadding=0  >"
    end if

	If Not rsfav.EOF = 0 Then


		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;FAVORITOS&nbsp;" & "&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"

		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"


	Else

		Do While Not rsfav.EOF

			CLAVE = rsfav(0)
			If xTrColor = "#FFFFFF" Then
				xTrColor = "#eeeeff"
			Else
				xTrColor = "#FFFFFF"
			End If

			XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=25 >"
			sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsfav("MSGOBJCOD") & "&Tipo=" & rsfav("MSGOBJTIP") & "&MAINAPL=" & xMAINAPL
			sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
			sHref = "&nbsp;"
			' response.write sActcon
			''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

			sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)

			If rsfav("ESTADO") & "" = "3" Then

				TITULOLNK = "Consultar Documento..."

			Else
				sFuncion = ""
				sHref = "#"
				TITULOLNK = "EL Documento NO encuentra vigente debe eliminarse de Favoritos..."
			End If

			sTipobjeto = EsqnomPx(rsfav("MSGOBJTIP"), rsfav("MSGOBJCOD"), ofv)



			wmsgx = sTipobjeto & ""
			wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



			sImagen = wNomL

			If xarc = "NO" Then
				xarc = "&nbsp;"
			End If

			sLink = sImagen

			XTSLNKM = ofv.GenLink("", "nt  title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, wmsgx, "si")

			sImage = ofv.GenImage("", "", "", "", "Imagenes/Botdelx.jpg", "0", "")
			vlk = "../novedmenu.asp?tab=view2"
			sHref = "asp/EliminarEN.asp?ID=" & CLAVE & "&Tabla=objmsg&volver=" & vlk
			sHref = sHref & "&Codobjeto=" & rsfav("MSGOBJCOD") & "&Tipobjeto=" & rsfav("MSGOBJTIP") & "&Eliminar=SI"

			CELDA = BotonEliminar("", sHref, "fav" & CLAVE,"N")

            '**** TITULO DE LAS AUTORIZACIONES ****'
			'**************************************'

			'***** Incluye el link al WKF ******'

			XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"
			XTablaN = XTablaN & "<td " & xxsclase2 & " colspan=8 ><font color='#000000' size='1.5'><b><div style=""z-index:99999;width:40%;margin:0 auto;border-radius:10px 10px 10px 10px;background-color:#dddddd;box-shadow:3px 0px #cccccc;height:30px;text-align:center;line-height:30px;cursor:pointer;"">"   &  celda & "&nbsp;" & XTSLNKM & "</div></b></font></TD>"

			XTablaN = XTablaN & "</TR>"

			'**************************************'
			'**** FIN TITULO DE LAS AUTORIZACIONES ****'
			'**************************************'



			'XTablaN = XTablaN & CELDA
'
			'XTablaN = XTablaN & "<td>"
			'XTablaN = XTablaN & XTSLNKM & "</TD>"

			'XTablaN = XTablaN & "</TR>"













			rsfav.movenext
		Loop


	End If




	XTablaN = XTablaN & "</table></div>"
	Favoritos = XTablaN

End Function


Function SUSCRIPCION(tit)
	Dim XTablaN
	'Strsql=" Select *,B.ESTADO   From objmsg a "
	'STRSQL =STRSQL  & " LEFT JOIN OBJPRTY B ON(A.MSGOBJCOD= B.CODOBJETO) "
	'strsql=strsql & " where a.MSGTIPOMESS=2  AND usr = '" & SESSION("USUARIO") & "'"

	xsclase = " CLASS=ietx "
	xsclase2 = "class=zzzzz "
	StrSql = "SELECT * "
	StrSql = StrSql & "FROM csnovedades   a inner join objprty c on(a.codobjeto=c.codobjeto and c.estado=3) "
	StrSql = StrSql & "inner join objmsg b on( a.codobjeto=b.MSGOBJCOD and b.MSGTIPOMESS=2 and  b.usr = '" & SESSION("USUARIO") & "')"

	StrSql = StrSql & " Where a.fecinicio > " & dFechai - 1
	StrSql = StrSql & " and a.fecinicio <  " & dFechai + 6
	StrSql = StrSql & " And a.tipo = 3 "
	'response.write strsql
	xTrColor = ""
	' vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID
	'response.write vuelta
	Set rsfav = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


    xsclase = " CLASS=ietx "
    xsclase2 = "class=istt "
	'''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "

	XTablaN ="<div id='DTNOVF' style='width:100%;xbackground-color:#dddddd;' "
	XTablaN = XTablaN & "xborder-style:solid;border-width:3;border-color:#ffffff;' >"




	    if tit= "S" then
	      '********* ENCABEZADO DE SECCIONES

		XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1     >"
		XTablaN = XTablaN & "<tr HEIGHT=25  >"
		''''XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(Capitalize(atit(xmenu)) & "", " ", "&nbsp;") &  "</TD>"
		XTablaN = XTablaN & "<td " & xsclase & "   >SUSCRIPCIONES</TD>"
		XTablaN = XTablaN & "</TR>"
		XTablaN = XTablaN & "</table>"

	    XTablaN = XTablaN & "<table  width=98% cellspacing=0 cellpadding=0  >"

	    else
	    	XTablaN = XTablaN & "<table  width=98% cellspacing=0 cellpadding=0  >"
    end if

	'XTablaN = XTablaN & "<table   width=98% cellspacing=0 cellpadding=0  >"


	If rsfav.EOF Then


		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;SUSCRIPCIONES&nbsp;" & "&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"

		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"
		xblink = ""

	Else

		xblink = " <script> var PgColor='#FF9B42'; function SetPgColor()"
		xblink = xblink & " {"
		xblink = xblink & " "

		xblink = xblink & " if (PgColor=='#FF9B42')"
		xblink = xblink & "{"
		xblink = xblink & "PgColor='#ffffff';document.getElementById('Xsus').style.backgroundColor='#ffffff';setTimeout('SetPgColor()', 500);"
		xblink = xblink & " }"
		xblink = xblink & " else"
		xblink = xblink & " {"
		xblink = xblink & " PgColor='#FF9B42';document.getElementById('Xsus').style.backgroundColor='#FF9B42';setTimeout('SetPgColor()', 500);"
		xblink = xblink & " } "

		xblink = xblink & " }"
		xblink = xblink & "setTimeout('SetPgColor()', 500);"
		xblink = xblink & "</script>"

		XTablaNdelim = ""  'XTablaNdelim & "</TR>"
		'     XTablaNdelim = XTablaNdelim & "<tr><td><hr>"
		'   XTablaNdelim = XTablaNdelim  & "</TD></tr>"


		Do While Not rsfav.EOF

			CLAVE = rsfav(0)
			If xTrColor = "#FFFFFF" Then
				xTrColor = "#eeeeff"
			Else
				xTrColor = "#FFFFFF"
			End If

			XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=25 >"
			sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsfav("MSGOBJCOD") & "&Tipo=" & rsfav("MSGOBJTIP") & "&MAINAPL=" & xMAINAPL
			sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
			sHref = "&nbsp;"

			sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)



			sTipobjeto = EsqnomPx(rsfav("MSGOBJTIP"), rsfav("MSGOBJCOD"), ofv)



			wmsgx = sTipobjeto & ""
			wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



			sImagen = wNomL

			If xarc = "NO" Then
				xarc = "&nbsp;"
			End If

			sLink = sImagen

			XTSLNKM = ofv.GenLink("", "nt  title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, wmsgx, "si")

			xtfechavig = fechaArg(CDate(rsfav(3) + CLng(CDate("01-01-1900"))))
			XTMSGSEND = "<font color='#AA0000' size=2 >&#0149;</font></B><B><font color=#000000 >" & xtfechavig & "</font></B>"
			XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
			wmsgx = Trim(rsfav(5) & "")

			If Len(wmsgx) < 5 Then
				wmsgx = EsqnomPx(rsfav(9), rsfav(8), ofv) '& " VIGENTE A PARTIR DEL " & fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900")))) & ". "
			Else
				'if xmenu=q3 then
				' wmsgx = UCase(wmsgx) ' & "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
				' else'response.write "aa" & len(wmsgx) & rsx(5)
				wmsgx = UCase(wmsgx)  '& "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
				' end if
			End If

			XTMSGSEND = XTMSGSEND & "<font color=#000000 >&nbsp;" & wmsgx & "</font>"
			IF  session("APLMDCLI")="BGAL" THEN
			  XTSLNKM = XTSLNKM  & XTMSGSEND
			  XTSLNKM =  ofv.GenLink("", "nt  title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, XTMSGSEND, "si")
			ELSE
			  XTSLNKM = XTSLNKM & "<br><br>" & XTMSGSEND
			  XTSLNKM = "<br>" & ofv.GenLink("", "nt  title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, XTMSGSEND, "si") '& "<br>"
			END IF

			sImage = ofv.GenImage("", "", "", "", "Imagenes/Botdelx.jpg", "0", "")
			vlk = "../novedmenu.asp?tab=view3"
			sHref = "asp/EliminarEN.asp?ID=" & CLAVE & "&Tabla=objmsg&volver=" & vlk
			sHref = sHref & "&Codobjeto=" & rsfav("MSGOBJCOD") & "&Tipobjeto=" & rsfav("MSGOBJTIP") & "&Eliminar=SI"

			CELDA = "" 'BotonEliminar("", sHref, "sus" &CLAVE)


			'XTablaN = XTablaN & CELDA
            '**************************************'

			'***** Incluye el link al WKF ******'

			XTablaN =  XTablaN & "<tr HEIGHT='30' >"
			XTablaN = XTablaN & "<td " & Xxsclase2 & " xcolspan=8 ><font color='#000000' size='1.5'><b><div style='width:40%;margin:0 auto;border-radius:10px 10px 10px 10px;background-color:#dddddd;box-shadow:3px 0px #cccccc;height:30px;text-align:center;line-height:30px;cursor:pointer;'>"   &  XTSLNKM & "</div></b></font></TD>"

			XTablaN = XTablaN & "</TR>"

'**************************************'
			'XTablaN = XTablaN & "<td " & xsclase2 & ">"
			'XTablaN = XTablaN & XTSLNKM & "<br><br></TD>"

			'XTablaNdelim = XTablaNdelim & "</TR>"






			rsfav.movenext
		Loop


	End If




	XTablaN = XTablaN & "</table></DIV>" & xblink





	SUSCRIPCION = XTablaN

End Function
Function ULTFORMCONS(tit)
dim fila
dim celda
	Fila    = ""



	xTrColor = "#FFFFFF"

	stitem = "Tema"
	sIcono = "ICorTema.jpg"
	sNuevoTipo = "SUBTEMAS"
	nNivel = 0

	sID = 0
	nv = 0
	cborderB = Session("CBRDBOT")
	cborderl = Session("CBRDLFT")
	borderB = ""
	borderl = ""



	XWHERE = " and USUARIO = '" & session("usuario") & "' "
	XWHERE = XWHERE & " AND b1.CODOBJETO > 0 "

	If wkaplic = "OR" Or wkaplic = "PL" Then

		XWHERE = XWHERE & " AND b1.TIPOBJETO = 'RSK' "

	Else

		XWHERE = XWHERE & " AND b1.TIPOBJETO <> 'RSK' "

	End If

	XORDER = " group by b1.codobjeto ORDER BY fecha desc "


	StrSql = ""
	cantreg=atit(5,3)
	StrSql = StrSql & " SELECT top " & cantreg & " b1.CODOBJETO, max(b1.TIPOBJETO), max(finicio) as fecha  "
	StrSql = StrSql & " FROM OBJALOGDET B1"
	StrSql = StrSql & " inner join  objprty b on (b1.CODOBJETO=b.codobjeto and b.estado =3) "
	StrSql = StrSql & " inner join objetos c on (b1.CODOBJETO=c.codobjeto and b1.tipobjeto=c.tipobjeto) "
	StrSql = StrSql & " inner join  TIPOBJETO Tb on (b.TIPOBJETO=Tb.TIPOBJETO and tb.tipocpt ='K') "

	'StrSql = StrSql & " WHERE TACC = 'CONSEN'  " & XWHERE
    if  session("UN")=1 then
	    StrSql = StrSql & " WHERE  TACC = 'CONSEN'  " & XWHERE
	 else
	    ' StrSql = StrSql &  "inner join objetos b2 on (b2.CODOBJETO=b.codobjeto) "
	     StrSql = StrSql & " WHERE C.objun=" & session("UN") & " and TACC = 'CONSEN'  " & XWHERE
	end if
	StrSql = StrSql & XORDER

	 ' response.write strsql
	Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If  rsx.eof Then

		Celda   = ""
		Celda   = Celda & ofv.GenCelda("","ttt","","2%","","","","&nbsp;","si")
		Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","","No hay elementos asociados","si")

		Fila    = fila & ofv.GenRow("","","","","25","","",Celda,"si")

	Else

		Celda   = ""

           if tit="S" then
            xsclase = " CLASS=ietx "
			'celda =  "<td " & xsclase & " style='text-align:left;background-color:#ffffff;box-shadow:#ffffff 0px 0px;' >" & "ULTIMOS FORMULARIOS CONSULTADOS" & "</TD>"

			if  session("APLMDCLI")="BGAL" THEN
			celda =  "<td " & xsclase & " style='text-align:left;background-color:#ffffff;box-shadow:#ffffff 0px 0px;font-family:calibri;font-size:16px;border-bottom-color:#000000;border-bottom-style:solid;border-radius:0px;border-width:1px;box-shadow:0px 0px #ffffff;' >" & "Mis Últimas Consultas" & "</TD>"
				Else
				celda =  "<td " & xsclase & " style='text-align:left;background-color:#ffffff;box-shadow:#ffffff 0px 0px;' >" & "ULTIMOS FORMULARIOS CONSULTADOS" & "</TD>"
			End if

			fila=fila &  ofv.GenRow("","aaa ","","","20","","",Celda,"si")
          end if
		Do Until rsx.eof


            if  session("APLMDCLI")="BGAL" THEN
             xTrColor = "#FFFFFF"
            ELSE
			  If xTrColor = "#FFFFFF" Then
				xTrColor = "#eeeeff"
			  Else
				xTrColor = "#FFFFFF"
			  End If
			END IF

			xarc = "" '"<font face=webdings color=#993333 size=4  >&#0157;</font>&nbsp"


			sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
			sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
			sHref = "&nbsp;"
			' response.write sActcon
			''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

			sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


			sTipobjeto = EsqnomPx(WKSPTIPOOBJ, rsx(0), ofv)






			wmsgx = xarc & sTipobjeto & ""
			wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



			sImagen = wNomL

			If xarc = "NO" Then
				xarc = "&nbsp;"
			End If

			sLink = sImagen

			XTSLNKM = ofv.GenLink("", "nt  title='Consultar Formulario...' ", sHref, "", "", sFuncion, wmsgx, "si")


			celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "' style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","","","","", XTSLNKM,"si")

			Fila    = Fila & ofv.GenRow("","aaa ","","","25","","",Celda,"si")


			RSx.MOVENEXT
		Loop

	End If



	ULTFORMCONS= ofv.GenTabla("","aaa   ","","100%","","0","","0","","","",Fila,"si")

    if tit="S" then
     XTablaN =  "<div id='DTNOV" & tit & "' style='width:100%;xbackground-color:#dddddd;' "
	 XTablaN = XTablaN & "xborder-style:solid;border-width:3;border-color:#ffffff;' >"
	 XTablaN = XTablaN & ULTFORMCONS & "</div>"
     ULTFORMCONS=XTablaN
    end if
	Call OFV.cerrarconsulta(rsx)

End Function



Public Function BotonEliminar(sNombre, sHref, CLAVE,dato)

	wktitdelete = "Confirma la eliminacion "
	sImageELs = "<font face='corflaticon' size=3 color='#bb0000'>&#0114;</font>"
	sformobj = "FRMOBJ" & CLAVE
	sAccion = sHref
	formobj = ofv.genform(sformobj, "AAA STYLE='display:none;' ", sAccion, "", "post", "", "", "si")
	sfunobj = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "
    if  Dato <> "" THEN
     BotonEliminar = "<span  title='" & wktitdelete & "' " & sfunobj & ">"  & sImageELs & "</span>"
    ELSE
	 BotonEliminar = ofv.GenCelda("", "ut WIDTH=2% style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ", "", "2", "", "", "", sImageELs, "si")
	END if

	xxform = xxform & formobj


End Function
Function armapedidos()


	Dim tabla1f

	SQL= " select top 5 a.fechaini,a.fechafin,a.objid,isnull(a.nombre,'') as nombre,a.solicitante,a.responsable,a.estado,b.estado AS EST_DESCRI,TIPOBJETO from cs_pedidos a "
	sql=sql & " inner join estadosform b on (b.objid=a.estado) where RESPONSABLE IS NOT NULL AND solicitante='" & session("usuario") & "' AND RESPONSABLE <> '" & session("usuario") & "'"
	sql=sql & " order by  fechaini desc,a.objid desc"
	'response.write strsql
	fi = CLng(CDate("01-01-1900"))
	Set rsx = ofv.crearconsultaEx(SQL,cn,1,parametros)
	If  rsx.eof Then

		armapedidos=""

	Else

		Celda   = ""
		tabla1f =    "<div id='DTTEMAPD' style='float: left;width:98%; " & Chr(13)
		tabla1f = tabla1f & "border-style:solid;border-width:2;border-color:#ffffff;' >"

		tabla1f = tabla1f & "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
		tabla1f = tabla1f & "<tr HEIGHT=25  >"
		xsclase = " CLASS=ietx "
		''''tabla1 = tabla1 & "<td " & xsclase & "  >" & Capitalize("INDICE NORMATIVO") & "</TD>"
		tabla1f = tabla1f & "<td " & xsclase & "  >" & "ULTIMAS SOLICITUDES" & "</TD>"
		tabla1f = tabla1f & "</TR>"
		tabla1f = tabla1f & "</table>" & Chr(13)

		''''height:400px;

		tabla1f = tabla1f & "<div id='DTTEMAINTFPD' style='overflow:auto;" & wstyl & ";"
		tabla1f = tabla1f & "border-style:solid;border-width:1;border-color:#dddddd;' >" & Chr(13)


		Fila    = ""



		xTrColor = "#FFFFFF"

		stitem = "Tema"
		sIcono = "ICorTema.jpg"
		sNuevoTipo = "SUBTEMAS"
		nNivel = 0

		sID = 0
		nv = 0
		cborderB = Session("CBRDBOT")
		cborderl = Session("CBRDLFT")
		borderB = ""
		borderl = ""





		Do Until rsx.eof



			If xTrColor = "#FFFFFF" Then
				xTrColor = "#eeeeff"
			Else
				xTrColor = "#FFFFFF"
			End If
			xarc = "<font face=webdings color=#993333 size=4  >¤</font>&nbsp"


			'sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
			' sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
			aspeditar = "crearpedido.asp?pedido=VERW&idpedido=" & Rsx("OBJID") & "&XPC=" &Rsx("OBJID")
			sHref = aspeditar '"javascript:void(0);"

			sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sHref & "'); return false;" & Chr(34)


			wmsgx= xarc
			If Len(RSX("NOMBRE")) > 60 Then
				wmsgx= wmsgx  & Left(Trim	(RSX("NOMBRE")) ,60) & "..." & " (<i>" & RSx("EST_DESCRI") & "</i>) <br>"
			ElseIf Len( Trim(RSX("NOMBRE")&"")) = 0 Then
				wmsgx=  wmsgx  & " <i>S/D</i>" &  " (<i>" & RSx("EST_DESCRI") & "</i>)&nbsp;<br>"
			Else
				wmsgx=  wmsgx & RSX("NOMBRE") & " (<i>" & RSx("EST_DESCRI") & "</i>)&nbsp;<br>"
			End If

			wmsgx=    wmsgx  & "<i>NRO: </i>" & Rsx("OBJID") &"&nbsp;&nbsp"
			XX=GETINFOUSER(rsx("responsable"))
			wmsgx= wmsgx & "<i>RESP: </i>" & SESSION("SGU_NOMBRE")
			'  xtfechaINI= fechaArg(CDate(RSX("FECHAINI") + CLng(CDate("01-01-1900"))))


			XTSLNKM =  ofv.GenLink("", "nt   title='VER SOLICITUD NRO : "  & Rsx("OBJID") & "'" , sHref, "", "", sFuncion, wmsgx, "si")


			celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "' style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","", XTSLNKM,"si")

			Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")


			RSx.MOVENEXT
		Loop


		tabla1f=  tabla1f & ofv.GenTabla("","aaa   ","","95% ","","0","","0","0","","",Fila,"si")

		tabla1f = tabla1f & "</DIV>"
		tabla1f = tabla1f & "</DIV>"  & Chr(13)
		armapedidos= tabla1f



	End If






	Call OFV.cerrarconsulta(rsx)

End Function

Function tieneMensaje(sCod)

	sql = "select mensaje from novedades where codigo='" & sCod & "'"
	Set rsnov = ofv.crearconsultaEx(sql, cn, 1, parametros)

	sTiene = 0
	If Not rsnov.eof Then

		If rsnov("mensaje") & "" <> "" Then sTiene = 1

	End If

	tieneMensaje = sTiene

End Function

Function noveformatEX(xmenu, XMENSEL,TITVA,base)


	XTablaN = ""

	acctitblks = 0
	acctitlin = 0
    xTrColor="#FFFFFF"

	'xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "

	xsclase = " CLASS=ietx "
	xsclase2 = "class=istt "
	'''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "




		XTablaN = XTablaN & "<div id='DTNOV" & xmenu & "' style='width:100%;background-color:#ffffff;' "


	'********* ENCABEZADO DE SECCIONES
	XTablaN = XTablaN & "xborder-style:solid;border-width:3;border-color:#ffffff;' >"
	XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1      >"
	XTablaN = XTablaN & "<tr HEIGHT=25  >"
	''''XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(Capitalize(atit(xmenu)) & "", " ", "&nbsp;") &  "</TD>"
      XSTYLE= ""
	If SESSION("APLMDCLI") = "BGAL" Then
	    IF XMENU=5 OR XMENU=11 THEN ' ultimos consultados
	     xsclase=" CLASS=ietxBG "
	     XSTYLE= " style='text-align:left;border-bottom-color:#000000;border-bottom-style:solid;border-radius:0px;border-width:1px;box-shadow:0px 0px #ffffff;'"
        END IF

	 XTablaN = XTablaN & "<td " & xsclase &  XSTYLE  & " >" & Replace(atit(xmenu,1) & "", " ", "&nbsp;") & "</TD>"

    Else

		XTablaN = XTablaN & "<td " & xsclase &  XSTYLE & "  >" & Replace(atit(xmenu,1) & "", " ", "&nbsp;") & "</TD>"
	End if

	XTablaN = XTablaN & "</TR>"
	XTablaN = XTablaN & "</table>"
	If TITVA = "S" Then
		XTablaN =XTablaN &    "<div id='DTNOVINT" & xmenu & "' style='"
	Else
	    XTablaN =    "<center><div id='DTNOVINT" & xmenu & "' style='"

	    If xmenu = 2 Then

				    XTABLAN = XTablaN & " width:50%;"

			Else

				   XTABLAN = XTablaN & ""

	     End If


	End If



	XTablaN = XTablaN & "xborder-style:solid;border-width:1;border-color:#ff0000;' >"

	XTablaN = XTablaN & "<table    width=99% cellspacing=0 cellpadding=0  >"


	If TNTabla = 0 Then


		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(Menu,1) &  "&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"

		XTablaN = XTablaN & "<tr>"
		XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
		XTablaN = XTablaN & "</TR>"


	Else
        'IF XMENU=5 THEN RESPONSE.WRITE TNTabla
		For INTabla = 1 To TNTabla

			'    acctitlin  = acctitlin + 2


			xtmcod = ANTabla(INTabla, 0)
			xtipom = ANTabla(INTabla, 1)
			xtpri = ANTabla(INTabla, 2)
			xtfvig = ANTabla(INTabla, 3)
			xthvig = ANTabla(INTabla, 4)
			xtsadj = ANTabla(INTabla, 5)
			xtslnk = ANTabla(INTabla, 6)
			xfromusr = ANTabla(INTabla, 7)
			xtmsg = ANTabla(INTabla, 8)
			xtmsgimg = ANTabla(INTabla, 9)
			xtmsglnk = ANTabla(INTabla, 10)
			xtmsgTYPE = ANTabla(INTabla, 11)



			xfclas = "dt"
			sImage = "&nbsp;"

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

			sglmod = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
			sImage = ofv.GenLink("", xfclas & " title='" & sgltit & "' ", "", "", "", "", sglmod, "si")

			If xtfvig = dFechai Then
				XMARQ = sglmar
			Else
				XMARQ = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
			End If

			If xtipom = 5 Then
				xtfechavig = xthvig
				XTCOUNT = InStr(xtfechavig, " ")
				If XTCOUNT > 0 Then
					xtfechavig = Left(xtfechavig, XTCOUNT)
					xtfechavig = GTFechadorsh(CLng(CDate(xtfechavig)) - CLng(CDate("01-01-1900")))
				End If
			Else
				xtfechavig = GTFechadorsh(xtfvig)
				XTCOUNT = InStr(xtfechavig, " ")
				If XTCOUNT > 0 Then xtfechavig = Left(xtfechavig, XTCOUNT)


			End If

			Xnlnk = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
			Xnadj = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

			If xtsadj = "S" Then Xnadj = sgladj
			If xtslnk = "S" Then Xnlnk = sgllnk









			xtmsgVRS = "0" & ANTabla(INTabla, 12)
			xtmsgVRS = CInt(xtmsgVRS)
			xtmsgTIT = ANTabla(INTabla, 13)
			xtmsgCOP = ANTabla(INTabla, 14)
			xtmsgimg = ANTabla(INTabla, 15)
			xtmsgVID = ANTabla(INTabla, 16)


			xtmsgVAENC = ANTabla(INTabla, 17)
			xtmsgVAFORO = ANTabla(INTabla, 18)
			xtmsgENCCOD = ANTabla(INTabla, 19)
			xtmsgFOROCOD = ANTabla(INTabla, 20)

			xtmsgdoccod = ANTabla(INTabla, 21)


			If xtmsgVRS = 2 Then

				acctitlin = acctitlin + 2

				'xtmsglnk    = replace(xtmsglnk,"XXMSGXX","<font STYLE='color:#990000;font-size:10pt;font-weight:bold;' >" & xtmsgTIT & "</font>")
				xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

				'  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
				'  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
				'  XTABLAN = XTablaN &  "</TR>"

				XTablaN = XTablaN & "<tr HEIGHT='20' >"
				XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
				For yy = 1 To 6
					XTablaN = XTablaN & "<td class=UT width='15%'  >&nbsp</TD>"
				Next
				XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
				XTablaN = XTablaN & "</TR>"

				''  TIPO DE NOVEDAD NOTICIA DEL DIA
				''INCIDENTE 1366
				If xtipom = "1" Or xmenu = "2" Then

					xSolasp = ""
					xSolasp = xSolasp & "novedmenuview.asp?xxmsg=" & CStr(xtmcod) & "&Pantalla=MAIN"




					XTablaN = XTablaN & "<form name='VNGP" & CStr(xtmcod) & "' id='VNGP" & CStr(xtmcod) & "' action='" & xSolasp & "' "
					XTablaN = XTablaN & " method=post target='fesquemas' style='display:none;' >"
					' XTablaN = XTablaN & " method=post target='_new' style='display:none;' >"
					XTablaN = XTablaN & "</form>"



					xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

					XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"


					If SESSION("APLMDCLI") = "BGAL" Then
						XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8   ><MARQUEE>" & xtmsglnk & "</MARQUE></a></TD>"

					Else

						XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8   >" & xtmsglnk & "</a></TD>"


					End If


					XTablaN = XTablaN & "</TR>"





					'' Otros tipos de novedad
				Else

					'**************************************'
					'**** TITULO DE LAS AUTORIZACIONES ****'
					'**************************************'

					'***** Incluye el link al WKF ******'

					XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"
					if session("APLMDCLI") ="BGAL" THEN
					 XTablaN = XTablaN & "<td " & xxsclase2 & " colspan=8 ><div style=""width:80%;margin:0 auto;border-radius:10px 10px 0 0;background-color:#dddddd;box-shadow:3px 0px #cccccc;height:auto;text-align:center;line-height:auto;cursor:pointer;"">" & xtmsglnk & "</div></TD>"
					ELSE
					 XTablaN = XTablaN & "<td " & xxsclase2 & " colspan=8 ><font color='#000000' size='1.5'><b><div style='z-index:9999;width:40%;margin:0 auto;border-radius:10px 10px 0 0;background-color:#dddddd;box-shadow:3px 0px #cccccc;height:auto;text-align:center;line-height:30px;cursor:pointer;>wwwww  " & xtmsglnk & "</div></b></font></TD>"

					END IF

					'XTablaN = XTablaN & "<td " & xxsclase2 & " colspan=8 ><font color='#000000' size='1.5'><b><div style='z-index:9999;width:40%;margin:0 auto;border-radius:10px 10px 0 0;background-color:#dddddd;box-shadow:3px 0px #cccccc;height:30px;text-align:center;line-height:30px;cursor:pointer;>" & xtmsglnk & "</div></b></font></TD>"

					XTablaN = XTablaN & "</TR>"

					'**************************************'
					'**** FIN TITULO DE LAS AUTORIZACIONES ****'
					'**************************************'



				End If


				If xtipom = "9" Then

					acctitlin = acctitlin + 2

					XTMSGSEND = ""
					XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP
					XTablaN = XTablaN & "<tr  >"
					if session("APLMDCLI")="BGAL" THEN
					 XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ></TD>"
					  XTablaN = XTablaN & "<td class=ut colspan=8 STYLE='color:#000000;text-align:justify' >"
					ELSE
					 XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B>"
					 XTablaN = XTablaN & "<font color='#AA0000' size=2 >&#0187;</font>"
					 XTablaN = XTablaN & "<font color='#0000AA' size=1 >&#0124;</font>"
					 XTablaN = XTablaN & "</B></TD>"
					 XTablaN = XTablaN & "<td class=ut colspan=8 STYLE='color:#000000;text-align:justify' >"
					END IF
					if session("APLMDCLI")="BGAL" THEN
					 XTablaN = XTablaN & "<div style=""border-right:0.5px solid;margin:0 auto;width:78%;xborder-color:red;Xborder-style:solid;xborder-width:1px;border-radius:0 0 0px 0px;padding:5px;Xbox-shadow:3px 3px #dddddd;text-align:left;""><br><font color='#AA0000' size=2 >&#0187;</font>" & XTMSGSEND  & "</div>"
					 XTablaN = XTablaN  & "</TD>"
					ELSE
					XTablaN = XTablaN & XTMSGSEND & "</TD>"
					END IF

					XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
					XTablaN = XTablaN & "</TR>"


					XTMSGSEND = ""
					XTMSGSEND = XTMSGSEND & "<font color=#888888 style='font-weight:bold;' size=1 >recomendado por " & ANTabla(INTabla, 7) & "</font>"

					XTablaN = XTablaN & "<tr HEIGHT='20' >"
					IF session("APLMDCLI")="BGAL" THEN
					 XTMSGSEND="<div style=""border-right:0.5px solid;background-color:#dddddd;margin:0 auto;width:79%;xborder-color:red;xborder-style:solid;xborder-width:1px;border-radius:0 0 0px 0px;padding:5px;Xbox-shadow:3px 3px #dddddd;text-align:left;""><br><font color='#AA0000' size=2 >&#0187;</font>" & XTMSGSEND  & "</div>"
					 XTablaN = XTablaN & "<td class=UT align='center' colspan=8 STYLE=';text-align:justify;Xbackground-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
					ELSE
					XTablaN = XTablaN & "<td class=UT align='center' colspan=8 STYLE=';text-align:justify;background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
					END IF




					XTablaN = XTablaN & "</TR>"

				End If


				'''''''''''''''''''''''''''''
				'XTablaN = XTablaN & "<tr HEIGHT='20' >"
				'XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
				'XTablaN = XTablaN & "</TR>"


				If xtmsgdoccod <> "0" Then



					StrSql = "Select * From objinfo where codobjeto = 0" & xtmsgdoccod & " "
					'response.write strsql
					Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs1.EOF Then
						'aca esta el error

						''''' AGREGADO PARA MEJORAS VERSION 4.6
						xtmsgCOP = xtmsgCOP   & "<BR>" & rs1(5) & "" '& "<br>" ' & rs1(6)
						If xtmsgCOP <> "" Then
							xtmsgCOP = xtmsgCOP & "<br><br>" & rs1(6)
						End If
						''''' AGREGADO PARA MEJORAS VERSION 4.6 FIN

						xtmsgCOP = xtmsgCOP & "<br><br>" & rs1(4) '& xtmsgdoccod
						xtmsgimg = rs1(7)
						xtmsgVID = rs1(9)

					End If

					StrSql = "Select * From objinfoIMG  where codobjeto = 0" & xtmsgdoccod & " ORDER BY objSECUEN, OBJID"

					wspimgdat = 6

				Else

					StrSql = "Select * From NOVEDADESIMG  where NOVEDAD = '" & xtmcod & "' ORDER BY SECUEN, OBJID"
					wspimgdat = 5
				End If



				If (xtmsgimg & "") = "1" Then

					'   response.write Strsql

					Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs1.EOF Then



						XTablaN = XTablaN & "<tr xbgcolor='#F0F0F0' >"
						XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

						XTablaN = XTablaN & "<table  cellspacing=0 cellpadding=1  width=100% >"
						XTablaN = XTablaN & "<tr   >"


						XSFSTYLE = " STYLE='display:block;' "

						Do While Not rs1.EOF

							acctitlin = acctitlin + 10

							xCelda = ""
							xCelda = xCelda & "<div id='DIMGN" & INTabla & "' style='margin-top:1;margin-left:1;"
							xCelda = xCelda & "width:90%;overflow:auto;"
							xCelda = xCelda & ""
							xCelda = xCelda & "' >" & Chr(13)
							xCelda = xCelda & "<IMG style='border-style:solid;border-width:1;border-color:#888888;' SRC=" & Chr(34) & rs1(wspimgdat) & Chr(34) & ">"
							xCelda = xCelda & "</div>" & Chr(13)

							XTablaN = XTablaN & "<td CLASS=UT align='center' " & XSFSTYLE & " width=100% >" & xCelda & "</TD>"
							XSFSTYLE = " STYLE='display:none;' "


							rs1.movenext
						Loop

						XTablaN = XTablaN & "</TR>"
						XTablaN = XTablaN & "</table>"

						XTablaN = XTablaN & "</TD>"
						XTablaN = XTablaN & "</TR>"


						XTablaN = XTablaN & "<tr HEIGHT='20' >"
						XTablaN = XTablaN & "<td class=UT colspan=4  >&nbsp;</TD>"
						XTablaN = XTablaN & "</TR>"


					End If






				ElseIf (xtmsgimg & "") = "2" Then

					acctitlin = acctitlin + 10


					XTablaN = XTablaN & "<tr bgcolor='#F0F0F0' >"
					XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
					XTablaN = XTablaN & "<a class=istt target='_blank' href='" & xtmsgVID & "'>" & xtmsgVID & "</a>"
					'  XTablaN = XTablaN & xtmsgVID
					'   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
					XTablaN = XTablaN & "</TD>"
					XTablaN = XTablaN & "</TR>"

					XTablaN = XTablaN & "<tr HEIGHT='20' >"
					XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
					XTablaN = XTablaN & "</TR>"

				End If

				acctitlin = acctitlin + 2

				'**** COPETE NOTICIA DEL DIA
				XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
				XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
				XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP
				'''''response.write xtmsgCOP & " xtmsgCOP - mensaje "

				XTablaN = XTablaN & "<tr>"
				'XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B><font color='#AA0000' size=2 >&#0149;</font></B></TD>"
				XTablaN = XTablaN & "<td class=ut  padding=0 colspan=8 STYLE=""border-spacing;text-align:justify;"" >"
				' LINK VER MAS
				''INC 1366
				'  If xtipom = "1" Then
				url = "novedmenuview.asp?xxmsg=" & xtmcod
				sFuncion = "  onclick=" & Chr(34) & "document.all.VNGP" & CStr(xtmcod) & ".submit(); WKlevanta('" & CStr(xtmcod) & "',this);return false; " & Chr(34) & " href=''"
				'sFuncion = "  onclick=" & Chr(34) & "alert('1');" & Chr(34)
				If xmenu = 2 Then
					xtmsglnk = ofv.GenLink("", "istt style='cursor:pointer ;' title='Ver Mas ...' ", "", "", "", sFuncion, "Ver Mas", "si")
				Else
					xtmsglnk = ""
				End If


				'***********************************************************************'
				'******* 			AUTORIZACIONES 								********'
				'******* Descripción que inclulye fecha, nombre del esquema  	********'
				'******* y paso para autorizar									********'
				'***********************************************************************'

				'XTablaN = XTablaN & "<div style=""margin:0 auto;width:39%;xborder-color:red;xborder-style:solid;xborder-width:1px;border-radius:0 0 10px 10px;padding:5px;box-shadow:3px 3px #dddddd;text-align:left;""><br>" & XTMSGSEND & sImagen1 & "<br><br>" & xtmsglnk & "</div>"
				if session("APLMDCLI") ="BGAL" THEN
                  XTablaN = XTablaN & "<div style=""border-right:0.5px solid;margin:0 auto;width:79%;xborder-color:red;xborder-style:solid;xborder-width:1px;border-radius:0 0 10px 10px;padding:5px;box-shadow:3px 3px #dddddd;text-align:left;"">" & XTMSGSEND & sImagen1 & "<br><br>" & xtmsglnk & "</div>"
                ELSE
                  XTablaN = XTablaN & "<div style=""margin:0 auto;width:39%;xborder-color:red;xborder-style:solid;xborder-width:1px;border-radius:0 0 10px 10px;padding:5px;box-shadow:3px 3px #dddddd;text-align:left;""><br>" & XTMSGSEND & sImagen1 & "<br><br>" & xtmsglnk & "</div>"
                END IF
				'******* FIN AUTORIZACIONES ********'


				'  Else
				'  XTablaN = XTablaN & XTMSGSEND & "</TD>"
				'  End If
				' XTablaN = XTablaN & "</TD>"



				''' ********
				'XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
				XTablaN = XTablaN & "</td></TR>"




				''''' PREGUNTA SI VA ENCUESTA '''''''''''''''''''''''

				'    xtmsgVAENC      = ANTabla(INTabla,17)
				'    xtmsgVAFORO     = ANTabla(INTabla,18)
				'    xtmsgENCCOD     = ANTabla(INTabla,19)
				'    xtmsgFOROCOD    = ANTabla(INTabla,20)



				If xtmsgVAENC = "S" Then

					acctitlin = acctitlin + 2

					StrSql = "Select * From emencuestas  where objid = 0" & xtmsgENCCOD
					Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs1.EOF Then

						'''''''xtmsgTITENC = Capitalize(rs1(17) & "")
						xtmsgTITENC = rs1(17) & ""
						xtmsgENCDER = "S"


						If (rs1(2) & "") = "1" Then
							StrSql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & xtmsgENCCOD
							StrSql = StrSql & " Order By BLKSECUEN,objid DESC"
							Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
							If Not rs1.EOF Then
								sblktvoto = rs1(4) & ""
								sblkcodigo = rs1(0) & ""
								If sblktvoto = "2" Then
									xtmsgENCDER = "N"
									''''''xtmsgTITENC = Capitalize(rs1(3) & "")
									xtmsgTITENC = rs1(3) & ""



									XTablaN = XTablaN & "<tr  >"
									XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' >&nbsp;</TD>"
									XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='color:#000000;' align='center' ><br>"

									XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=1  style='background-color:#eeeeee;border-style:solid;border-width:1;border-color:#dddddd;' >"

									XTablaN = XTablaN & "<tr HEIGHT=20  >"
									XTablaN = XTablaN & "<td class=vtI COLSPAN=3  >Encuesta</TD>"
									XTablaN = XTablaN & "</TR>"


									XTablaN = XTablaN & "<tr HEIGHT=20  >"
									XTablaN = XTablaN & "<td class=utI width=85% style='font-size:8pt;font-weight:normal;' >" & xtmsgTITENC & "</TD>"

									WSPDIVN = "DENC" & xtmsgENCCOD

									StrSql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & xtmsgENCCOD
									StrSql2 = StrSql2 & " and BLKcodigo = 0" & sblkcodigo
									StrSql2 = StrSql2 & " Order By ITMSECUEN,objid DESC"

									Set rs2 = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)
									Do While (Not rs2.EOF)

										acctitlin = acctitlin + 2


										WSPFRMN = "FENC" & xtmsgENCCOD & "IT" & rs2(0)
										WSPFRMA = "EMENCUESTASACT.ASP?TP=ES&OP=ACT&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
										WSPFRMR = "EMENCUESTASACT.ASP?TP=ES&OP=RES&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
										XTablaN = XTablaN & "<td class=ut width=5% align='center' >"
										WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"
										XTablaN = XTablaN & "<input type='button' style='width:45px;background-color:#cc0000;color:#ffffff;font-weight:bold;border-style:solid;border-width:1;border-color:#888888;' "
										XTablaN = XTablaN & "  value='&nbsp;" & rs2(4) & "&nbsp;' "
										XTablaN = XTablaN & "  onclick='"
										'  XTABLAN = XTablaN &  "alert("" voto:" & rs2(4) & " "");"
										'  XTABLAN = XTablaN &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
										XTablaN = XTablaN & "document." & WSPFRMN & ".submit();"
										XTablaN = XTablaN & "document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";"
										XTablaN = XTablaN & "' >"
										XTablaN = XTablaN & "</TD>" & WSPFORME
										rs2.movenext
									Loop

									acctitlin = acctitlin + 2

									xCelda = ""
									xCelda = xCelda & "<div id='" & WSPDIVN & "' style='z-index:1;margin-top:1;margin-left:1;display:none;"
									xCelda = xCelda & "width:99%;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
									xCelda = xCelda & "background-color:#F9F9F9;"
									xCelda = xCelda & "' >" & Chr(13)
									xCelda = xCelda & "&nbsp;"
									xCelda = xCelda & "</div>" & Chr(13)


									XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"
									XTablaN = XTablaN & "</TR>"

									XTablaN = XTablaN & "<tr HEIGHT=20  >"
									XTablaN = XTablaN & "<td class=utI width=85%  >&nbsp;</TD>"
									XTablaN = XTablaN & "<td class=vt COLSPAN=2 ALIGN='CENTER' ><A CLASS=NT HREF='" & WSPFRMR & "' TARGET='fesquemas' "
									XTablaN = XTablaN & " ONCLICK='document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";' "
									XTablaN = XTablaN & "><B>Resultados</B></TD>"
									XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"

									XTablaN = XTablaN & "</TR>"

									XTablaN = XTablaN & "<tr   >"
									XTablaN = XTablaN & "<td class=vtI COLSPAN=4 >&nbsp;</TD>"
									XTablaN = XTablaN & "</TR>"



									XTablaN = XTablaN & "</table>" & xCelda
									XTablaN = XTablaN & "</TD>"
									XTablaN = XTablaN & "<td class=ut   >&nbsp;</TD>"
									XTablaN = XTablaN & "</TR>"





								End If

							End If


						End If





					End If

				End If


				''''''''''''''''''''''''''''

				'XTablaN = XTablaN & "<tr HEIGHT='20' >"
				'XTablaN = XTablaN & "<td class=UT colspan=8 STYLE='border-radius:10px;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;box-shadow:5px 5px #dddddd;' >&nbsp;" & "aaaaaa" & "</TD>"
				'XTablaN = XTablaN & "</TR>"











			Else

				acctitlin = acctitlin + 1


				''''  xtmsg    = replace(ltrim(xtmsg & "")," ","&nbsp;")
				xtmsg = LTrim(xtmsg & "")
				xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsg & "")


				If xtmcod = "0" Then

					XTMSGSEND = "<font color=#000000 >" & xtmsglnk & "</font>"

				Else

					XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
					XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
					XTMSGSEND = XTMSGSEND & "<font color=#000000 >&nbsp;" & xtmsglnk & "</font>"

				End If

                if ANTabla(INTabla, 1) <> 13  then

				XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " HEIGHT='25' >"
				XTablaN = XTablaN & "<td class=" & xfclas & " align='center' >" & xtmsgimg & "</TD>"

				IF (xmenu=5 or xmenu=11) and session("APLMDCLI")="BGAL" THEN

				 XTablaN = XTablaN & "<td class=uti colspan=6 STYLE='text-align:justify;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >" & XTMSGSEND & "</TD>"
				ELSE
				 XTablaN = XTablaN & "<td class=uti colspan=6 STYLE='text-align:justify;' >" & XTMSGSEND & "</TD>"
				END IF
				XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
				XTablaN = XTablaN & "</TR>"

               end if
				If xtmcod <> "0" Then

					' acctitlin  = acctitlin + 1


					XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=5 >"
					XTablaN = XTablaN & "<td class=ut colspan=8 valign=top align=right style='border-top-style:solid;border-top-width:1;border-top-color:#eeeeee;' >&nbsp;</TD>"
					XTablaN = XTablaN & "</TR>"

				End If


			End If

           IF  (xmenu=5 or xmenu=11)  and session("APLMDCLI")="BGAL" THEN
           xTrColor="#FFFFFF"
           ELSE
			If xTrColor = "#FFFFFF" Then
				xTrColor = "#eeeeff"
			Else
				xTrColor = "#FFFFFF"
			End If
           END IF

			'  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
			'  XTABLAN = XTablaN &  "<td class=UT colspan=8  ><HR  size=1 style='width:400;color:#ff0000;' >&nbsp;</TD>"
			'  XTABLAN = XTablaN &  "</TR>"




		Next


	End If
    IF  (xmenu=5 or xmenu=11) and session("APLMDCLI")="BGAL" THEN
	  XTablaN = XTablaN & "<tr HEIGHT='20' >"
	   XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
	  XTablaN = XTablaN & "</TR>"
    end if


	XTablaN = XTablaN & "</table>"
	'XTablaN = XTablaN & "<xHR  size=1 style='width:350;color:#eeeeee;' >"
	XTablaN = XTablaN & "</DIV>"


	XTablaN = XTablaN & "</DIV></center>"



	If TNTabla = 0 Then
		noveformatEX = ""
	Else

		If xmenu = 4 Then atmensajes = TNTabla

		ttitsecc = ttitsecc + 1

		'''''atitsecciones(ttitsecc, 0) = Capitalize(atit(xmenu))
		' atitsecciones(ttitsecc, 0) = atit(xmenu,1)
		' atitsecciones(ttitsecc, 1) = TNTabla
		' atitsecciones(ttitsecc, 2) = acctitlin + 2
		' atitsecciones(ttitsecc, 3) = XTablaN
		' atitsecciones(ttitsecc, 4) = 0
		' atitsecciones(ttitsecc, 5) = 0
		' IF xmenu & "" <> "4" THEN
		noveformatEX = XTablaN
		'END IF

	End If


End Function


Public Function BotonEliminarEx(sNombre, sHref,CLAVE)

 wktitdelete="Confirma la eliminacion "
 sImageELs = "<font face='corflaticon' size=3 color='#bb0000'>&#0114;</font>"
 sformobj   = "FRMOBJ" & CLAVE
 sAccion    = sHref
 formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
 sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "
 BotonEliminarEx="<span title='" & wktitdelete & "' " & sfunobj  & ">" &  sImageELs & "</span>"
 xxform=xxform & formobj


End Function
%>



