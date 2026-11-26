<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else

	''''''''''''
	Generar = ""

	volver = OFV.vueltaasp2("../")
	volver2 = OFV.vueltaasp2("")
'response.write Request.querystring
	sTipo = Request.querystring("Tipo")
	sNombre = Request.querystring("Nombre")
	sNombre = OFV.convertircar(sNombre, "_", " ")
	If Request.querystring("Menu") <> "" Then Session("Menu") = Request.querystring("Menu")

	sID = Request.querystring("ID")
	If sNombre = "" Then sNombre = UCase(sTipo)

	sitem    = OFV.convertircar(Request.querystring("sitem"), "_", " ")
	stitem   = OFV.convertircar(Request.querystring("stitem"), "_", " ")
	sticono  = Request.querystring("sticono")
	stnivel  = Request.querystring("stnivel")
	stestado = Request.querystring("stestado")
	CantRegMover=clng("0" & Request.querystring("CantRegMover"))
	CantRegMoversb=clng("0" & Request.querystring("CantRegMoversb"))

	CantregMostrar=10
	If Request.querystring("stesq") = "" Then
		stesq  = "TEMA"
	Else
		stesq  = Request.querystring("stesq")
	End If

	storig = "no"


	If Session("CorApli") = "COREM20" Then
		wkaplic = "EM"
	ElseIf Session("CorApli") = "COREP20" Then
		wkaplic = "PL"
	Else
		wkaplic = "OR"
	End If

	sbuscar=request.form("buscar")
	xbt=Request.querystring("xbt")
	' response.write "xbt=" & xbt
	If xbt ="" Then session("sbuscar")=""
	If request.form.count > 0  Then
		session("sbuscar")=sbuscar
	Else
		sbuscar=session("sbuscar")
	End If
	XXXX = NEWEMPRE("NO")

	If Session("sEstadon") <> "3" Then
		sfont    = "<font id='XXXX' color=#336699 >"
		sffont   = "</font>"
		sfont2   = "<font id='XXXX' color=#993333 >"
		sffont2  = "</font>"
		sfont3   = "<font id='XXXX' color=#777777 >"
		sffont3  = "</font>"
		sfont4   = "<font id='XXXX' >"
		sffont4  = "</font>"
		sfont5   = "<font id='XXXX' color=#FF00FF >"
		sffont5  = "</font>"
		sfont6   = "<font id='XXXX' color=#008800 >"
		sffont6  = "</font>"

	Else
		sfont    = "<font id='XXXX' >"
		sffont   = "</font>"
		sfont4   = "<font id='XXXX' >"
		sffont4  = "</font>"
	End If

	If Session("EMPMUN") = "S" Then
		WMATRMEMP    = "S"
		WMATRMEMPACT = "S"
		'ver aca 4/9/2020 ale
		'IF ATREMP(0,3) = 0
		'response.write SESSION("ZZZ")
		If ATREMP(0,2) = 0 Then
			WMATRMEMPACT = "N"
		End If

	Else
		WMATRMEMP = "N"
	End If

	SESSION("WMATRMEMP")    = WMATRMEMP
	SESSION("WMATRMEMPACT") = WMATRMEMPACT

	''  RESPONSE.WRITE "E " & ATREMP(0,0) & ATREMP(0,4) & " V" & ATREMP(0,2) & " A" & ATREMP(0,3) & "<BR>"


	If Session("Form") = "TEMAS-SUBTEMAS" Then
		Session("Form2") = OFV.convertircar(storig, "_", " ")
	Else
		Session("Form2") = Session("Form")
	End If

	volver2 = volver2 & "?Tipo=" & sTipo & "&Nombre=" & sNombre & "&ID=" & sID & "&Menu="
	volver2 = volver2 & Session("Menu") & "&sitem=" & OFV.convertircar(sitem, " ", "_")
	volver2 = volver2 & "&stitem=" & OFV.convertircar(stitem, " ", "_") & "&sticono=" & sticono
	volver2 = volver2 & "&stnivel=" & stnivel & "&stesq=" & stesq & "&stestado=" & stestado
	If Request.querystring("storig") <> "" Then
		volver2 = volver2 & "&storig=" & Request.querystring("storig")
	End If
	Session("ASP") = volver2

	borderB = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("APLSHW") & ";"
	borderl = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("APLSHW") & ";"

	'''  Session("CBRDICO") = "border-style:solid;border-width:1;border-color:" & Session("APLSHW") & ";"

	Session("CBRDICO") = "border-style:none;"


	xwbody = "bgcolor=#ccbbaa"
	If BrowserName() = "IE" Or BrowserName() = "CHROME" Then xwbody = " bgcolor=#dddddd "






	'sImagenM   = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
	'sImagenI2  = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
	'sImagenI3  = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	'sImagenEM = "<B><FONT FACE=CorFlaticon COLOR=#F7AA22 SIZE=3>&#48661;</FONT></B>"
	'  sImagen4si = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>&#052;</FONT></B>"
	'sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>&#57367;</FONT></B>"


	sImagenM = "<B><FONT FACE=""CorFlaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
	sImagenI2 = "<B><FONT FACE=""CorFlaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
	sImagenI3 = "<B><FONT FACE=""CorFlaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	sImagenEM = "<B><FONT FACE=CorFlaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
	sImagen4si = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>4</FONT></B>"
	sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"



	sHTML = "<html><body onload='window.top.ldfobj();' >"


	sHref = ""
	Fila = ""
	sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
	Celda = ""
	Celda = Celda & ofv.GenCelda("","VT colspan=8 ","","","","","",sImage & "&nbsp;<b><u>" & ATREMP(0,4) & "</u></b>","si")
	' Fila = fila & ofv.GenRow("","","top","","","","",Celda,"si")

	sImage = sImagenM

	Celda = ""
	Celda = Celda & ofv.GenCelda("","VT colspan=4 ","","30%","","","","&nbsp;" & sImagenM & "&nbsp;<b>" & UCASE(Session("FORMN")) & "</b>","si")
	Celda = Celda & ofv.GenCelda("","VT colspan=4 ","RIGHT","70%","","","",sImagenI2 & "&nbsp;<b>" & SESSION("SEGUSUARIO") & "</b>","si")

	Fila = fila & ofv.GenRow("","","top","","25","","",Celda,"si")

	Celda = ofv.GenCelda("","VT colspan=8 STYLE='BORDER-TOP-STYLE:RIDGE;BORDER-TOP-COLOR:#F7AA22;BORDER-TOP-WIDTH:2;' ","","","","","","&nbsp;","si")
	Fila = fila & ofv.GenRow("","","top","","","","",Celda,"si")


	Tabla = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")

	' sHTML = sHTML & Tabla

	Set cn = OFV.conectar(OFV.strconn2)
	nv = 0
	sHTML = sHTML & "<table cellpadding=0 cellspacing=0 width=90% style='margin-top:4;margin-left:4;' >"





	Select Case UCase(sTipo)
		Case "SUBTEMAS"
		If sID = "" Then sID = Session("sID")
		Session("XtemaSid")     = sID
		Session("XtemaSitem")   = sitem
		Session("XtemaStitem")  = stitem
		Session("XtemaSticono") = sticono
		Session("XtemaStnivel") = stnivel
		Session("XtemaStesq")   = stesq
		Session("Xnivel0mover")= Cantregmover


		sxlink = ""
		sxlink = sxlink & "<A class=nt HREF=Objetos.asp?Tipo=" & sTipo & "&ID="
		sxlink = sxlink & Session("XtemaSid") & "&sitem="
		sxlink = sxlink & OFV.convertircar(Session("XtemaSitem"), " ", "_")  &"&CANTREGMOVER="& Session("Xnivel0mover") &"&stitem="
		sxlink = sxlink & OFV.convertircar(Session("XtemaStitem"), " ", "_") & "&sticono="
		sxlink = sxlink & Session("XtemaSticono") & "&stnivel=" & Session("XtemaStnivel")
		sxlink = sxlink & "&stesq=" & Session("XtemaStesq") & " TARGET=fesquemas>"
		sxlinksig = ""
		sxlinksig = sxlinksig & "<A id=lnksig style='display:none;' class=nt HREF=Objetos.asp?Tipo=" & sTipo & "&ID="
		sxlinksig = sxlinksig & Session("XtemaSid") & "&sitem="
		sxlinksig = sxlinksig & OFV.convertircar(Session("XtemaSitem"), " ", "_")  &"&CANTREGMOVER="& Session("Xnivel0mover") + cantregmostrar  &"&stitem="
		sxlinksig = sxlinksig & OFV.convertircar(Session("XtemaStitem"), " ", "_") & "&sticono="
		sxlinksig = sxlinksig & Session("XtemaSticono") & "&stnivel=" & Session("XtemaStnivel")
		sxlinksig = sxlinksig & "&stesq=" & Session("XtemaStesq")
		sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face=Corflaticon >&#052;</font></a>"


		sxlinkant = ""
		sxlinkant = sxlinkant & "<A id=lnkant style='display:none;' class=nt HREF=Objetos.asp?Tipo=" & sTipo & "&ID="
		sxlinkant = sxlinkant & Session("XtemaSid") & "&sitem="
		sxlinkant = sxlinkant & OFV.convertircar(Session("XtemaSitem"), " ", "_")  &"&CANTREGMOVER="& Session("Xnivel0mover") - cantregmostrar  &"&stitem="
		sxlinkant = sxlinkant & OFV.convertircar(Session("XtemaStitem"), " ", "_") & "&sticono="
		sxlinkant = sxlinkant & Session("XtemaSticono") & "&stnivel=" & Session("XtemaStnivel")
		sxlinkant = sxlinkant & "&stesq=" & Session("XtemaStesq")
		sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face=Corflaticon >&#051;</font></a>"
		botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

		sxicon  = "ICorAbierto.jpg"
		sxiconv = "<font face='CorFlaticon' size=3 color='#f7aa22'>&#052;</font>"
		sxsid = sID
		If Trim(sID) = "xxx" Then
			sxicon = "ICorRecycle.jpg"
			sxiconv = "<font face='Corflaticon' size=3 color='#000000'>&#052;</font>"
			sxsid = 0
		End If


		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td class=ttt colspan=2 rowspan=2 style='" & Session("CBRDICO") & "' >"
		sHTML = sHTML & sxlink

		sHTML = sHTML & sxiconv
		''   sHTML = sHTML & "<img src=Iconos\" & sxicon& " border=0>"

		sHTML = sHTML & "</a></td>"
		sHTML = sHTML & "<td>&nbsp;</td>"
		sHTML = sHTML & "<td colspan=22 rowspan=2 >" & sxlink & Session("XtemaSitem") &"</a>" & xbotones& "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>&nbsp;</td>"
		sHTML = sHTML & "</tr>"


		StrSql = "Select Prodserv.codprodserv,Prodserv.descri,Prodserv.marcaps,Subtema.descri,"
		StrSql = StrSql & "Subtema.archivo,Prodserv.segaceso,Prodserv.segactual,Prodserv.PSTIPO ,Prodserv.PSURL "
		StrSql = StrSql & " From Prodserv, Subtema where "
		StrSql = StrSql & "Prodserv.marcaps = Subtema.tipobjeto and codtema = " & sxsid
		StrSql = StrSql & " order by Prodserv.indb, Prodserv.descri "

		sNuevoTipo = "ESQUEMAS"
		nNivel = 1
		Session("sID") = sID
		Session("EsqSup") = "Y"

		Case "ESQUEMAS"
		If stnivel = 1 Then
			Session("XstemaSid")     = sID
			Session("XstemaSitem")   = sitem
			Session("XstemaStitem")  = stitem
			Session("XstemaSticono") = sticono
			Session("XstemaStnivel") = stnivel
			Session("XstemaStesq")   = stesq
			Session("XstemaStest")   = "N"
			Session("XtemaStmover") =  Cantregmover
			Stsup                    = "Y"
		Else
			sesn = "X" & stnivel
			Session(sesn & "Sid")     = sID
			Session(sesn & "Sitem")   = sitem
			Session(sesn & "Stitem")  = stitem
			Session(sesn & "Sticono") = sticono
			Session(sesn & "Stnivel") = stnivel
			Session(sesn & "Stesq")   = stesq
			Session(sesn & "Stest")   = stestado
			Session(sesn & "Stestmover") = Cantregmover
			Stsup                     = "N"
		End If

		If Session("xorigen") = "tema" Then
			sxlink = ""
			sxlink = sxlink & "<A class=nt HREF=Objetos.asp?Tipo=SUBTEMAS&ID=" & Session("XtemaSid")
			sxlink = sxlink & "&sitem=" & OFV.convertircar(Session("XtemaSitem"), " ", "_")
			sxlink = sxlink & "&stitem=" & OFV.convertircar(Session("XtemaStitem"), " ", "_")
			sxlink = sxlink & "&sticono=" & Session("XtemaSticono") & "&stnivel="
			sxlink = sxlink & Session("XtemaStnivel") & "&stesq=" & Session("XtemaStesq")
			sxlink = sxlink & " TARGET=fesquemas>"

			sxicon = "ICorAbierto.jpg"
			If Trim(sID) = "0" Then sxicon = "ICorRecycle.jpg"

			sxiconv = "<font face='CorFlaticon' size=3 color='#f7aa22'>&#052;</font>"

			If Trim(sID) = "0" Then
				sxiconv = "<font face='CorFlaticon' size=3 color='#000000'>&#052;</font>"
			End If


			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td class=ttt colspan=2 rowspan=2 style='" & Session("CBRDICO") & "' >"
			sHTML = sHTML & sxlink
			sHTML = sHTML & sxiconv
			sHTML = sHTML & "</a></td>"

			'' & "<img src=Iconos\" & sxicon
			'''      sHTML = sHTML & " border=0></a></td>"


			sHTML = sHTML & "<td>&nbsp;</td>"
			sHTML = sHTML & "<td colspan=22 rowspan=2 >" & sxlink & Session("XtemaSitem") & "</a></td>"
			sHTML = sHTML & "</tr>"

			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>&nbsp;</td>"
			sHTML = sHTML & "</tr>"


		End If

		sxlink = ""
		sxlink = sxlink & "<A class=nt HREF=Objetos.asp?Tipo=ESQUEMAS&ID=" & Session("XstemaSid")&"&CANTREGMOVER="& Session("XtemaStmover")
		sxlink = sxlink & "&sitem=" & OFV.convertircar(Session("XstemaSitem"), " ", "_") & "&stitem="
		sxlink = sxlink & OFV.convertircar(Session("XstemaStitem"), " ", "_") & "&sticono="
		sxlink = sxlink & Session("XstemaSticono") & "&stnivel=" & Session("XstemaStnivel")
		sxlink = sxlink & "&stesq=" & Session("XstemaStesq")
		sxlink = sxlink & " ID=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		sxlink = sxlink & "-" & Session("XstemaStest")
		sxlink = sxlink & " name=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		sxlink = sxlink & "-" & Session("XstemaStest")
		sxlink = sxlink & " TARGET=fesquemas>"


		sHTML = sHTML & "<tr>"
		smas = "<font color=#993333 size=2><b>-</b></font>"
		sHTML = sHTML & "<td    >&nbsp;</td>"
		sHTML = sHTML & "<td style='" & borderl & borderb & "'>&nbsp;</td>"

		If BrowserName() <> "IE" Or BrowserName() <> "CHROME" Then
			sHTML = sHTML & "<img ID=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
			sHTML = sHTML & "-" & Session("XstemaStest")
			sHTML = sHTML & " name=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
			sHTML = sHTML & "-" & Session("XstemaStest")
			sHTML = sHTML & " border=0 src=Iconos\ICorAbre.jpg>"
		End If

		sxicon = "ICorAbierto.jpg"
		If Trim(sID) = "0" Then sxicon = "ICorRecycle.jpg"

		snid = Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		snid = snid & "-" & Session("XstemaStest")



		sxiconv = "<font id='" & snid & "' face='CorFlaticon' size=3 color='#f7aa22' onclick=""Editarx(this.id);"" >&#052;</font>"

		If Trim(sID) = "0" Then
			sxiconv = "<font id='" & snid & "' face='CorFlaticon' size=3 color='#000000' >&#052;</font>"
		End If

		sHTML = sHTML & "<td class=ttt colspan=2  rowspan=2 style='" & Session("CBRDICO") & "' >"

		sHTML = sHTML & sxlink
		sHTML = sHTML & sxiconv
		sHTML = sHTML & "</a></td>"
		' BOTON BUSQUEDA POR TEXTO BPN
		If session("APLMDCLI")="BPN" Then

			tablab =  "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"

			'''  sInputT  = ofv.GenerarInput("textot","","agregar","","80","dt style='border-style:solid;border-width:1;border-color:#666666;' ","")

			sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value='"& sbuscar & "' size='40' style='border-style:solid;border-width:1;border-color:#666666;' >"

			CELDA = ""
			CELDA = CELDA & ofv.GenCelda("", "UT  ", "right", "30%", "", "", "", sInputT, "si")

			sInput = ofv.GenerarInput("", "Buscar", "submit", "", "5", "at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;' ", "")
			CELDA = CELDA & ofv.GenCelda("", "vt", "", "70%", "", "", "", sInput, "si")

			formx = ofv.GenRow("", "", "", "", "", "", "", CELDA, "si")

			' Session(sesn & "Sid")     = sID
			' Session(sesn & "Sitem")   = sitem
			' Session(sesn & "Stitem")  = stitem
			' Session(sesn & "Sticono") = sticono
			' Session(sesn & "Stnivel") = stnivel
			' Session(sesn & "Stesq")   = stesq
			' Session(sesn & "Stest")   = stestado

			If  stnivel > 1 Then
				sTarget ="fesquemas"
				saccion="Objetos.asp?Tipo=ESQUEMAS&ID=" & Session(sesn & "Sid")
				saccion = saccion & "&sitem=" & OFV.convertircar(Session(sesn & "Sitem"), " ", "_")
				saccion = saccion & "&CantRegMover=0"
				saccion = saccion & "&stitem=" & OFV.convertircar(Session(sesn & "Stitem"), " ", "_") & "&sticono="
				saccion = saccion & Session(sesn & "Sticono") & "&stnivel=" & Session(sesn & "Stnivel")
				saccion = saccion & "&stesq=" & Session(sesn & "Stesq") & "&xbt=S"
			Else

				sTarget = "fesquemas"


				saccion="Objetos.asp?Tipo=ESQUEMAS&ID=" &  Session("XstemaSid")
				saccion = saccion & "&sitem=" & OFV.convertircar(session("XstemaSitem") , " ", "_")
				saccion = saccion & "&CantRegMover=0"
				saccion = saccion & "&stitem=" & OFV.convertircar( Session("XstemaStitem"), " ", "_") & "&sticono="
				saccion = saccion & Session("XstemaSticono") & "&stnivel=" & Session("XstemaStnivel")
				saccion = saccion & "&stesq=" &Session("XstemaStesq") & "&xbt=S"
			End If


			xform =  ofv.GenTabla("", "aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "50%", "", "0", "", "-2", "0", "", "", formx, "si")
			Fila = ofv.genform("Subconsulta", "", sAccion, sTarget, "post", "", xform, "si")

			tablab = tablab & Fila & "</DIV>" & Chr(13)
		Else
			tablab=""
		End If


		'  sHTML = sHTML & sxlink & "<img border=0 src=Iconos\" & sxicon
		'  sHTML = sHTML & " ID=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		'  sHTML = sHTML & "-" & Session("XstemaStest")
		'  sHTML = sHTML & " name=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		'  sHTML = sHTML & "-" & Session("XstemaStest") & "></a></td>"
		sxlinksig = ""
		sxlinksig = sxlinksig & "<A id=lnksig class=nt style='display:none;' HREF=Objetos.asp?xbt=" & xbt & "&Tipo=ESQUEMAS&ID=" & Session("XstemaSid")
		sxlinksig = sxlinksig & "&sitem=" & OFV.convertircar(Session("XstemaSitem"), " ", "_")
		' if CantRegMoversb=1 then
		'sxlinksig = sxlinksig & "&CantRegMover=" & CantRegMover
		' else
		sxlinksig = sxlinksig & "&CantRegMover=" & Session("XtemaStmover") + CantregMostrar
		' end if
		sxlinksig = sxlinksig & "&stitem=" & OFV.convertircar(Session("XstemaStitem"), " ", "_") & "&sticono="
		sxlinksig = sxlinksig & Session("XstemaSticono") & "&stnivel=" & Session("XstemaStnivel")
		sxlinksig = sxlinksig & "&stesq=" & Session("XstemaStesq")
		sxlinksig = sxlinksig & " ID=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		sxlinksig = sxlinksig & "-" & Session("XstemaStest")
		sxlinksig = sxlinksig & " name=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		sxlinksig = sxlinksig & "-" & Session("XstemaStest")
		' sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face=corflaticon>&#232;</font></a>"
		sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"

		sxlinkant = ""
		sxlinkant = sxlinkant & "<A id=lnkant class=nt style='display:none;' HREF=Objetos.asp?xbt=" & xbt & "&Tipo=ESQUEMAS&ID=" & Session("XstemaSid")
		sxlinkant = sxlinkant & "&sitem=" & OFV.convertircar(Session("XstemaSitem"), " ", "_")
		If  CantRegMover <= 0 Then
			sxlinkant = sxlinkant & "&CantRegMover=0"
		Else
			'if stnivel>2 then
			'sxlinkant = sxlinkant & "&CantRegMover="  &   CantRegMover
			'else

			sxlinkant = sxlinkant & "&CantRegMover="  &   Session("XtemaStmover") - CantregMostrar
			'end if
		End If
		sxlinkant = sxlinkant & "&stitem=" & OFV.convertircar(Session("XstemaStitem"), " ", "_") & "&sticono="
		sxlinkant = sxlinkant & Session("XstemaSticono") & "&stnivel=" & Session("XstemaStnivel")
		sxlinkant = sxlinkant & "&stesq=" & Session("XstemaStesq")
		sxlinkant = sxlinkant & " ID=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		sxlinkant = sxlinkant & "-" & Session("XstemaStest")
		sxlinkant = sxlinkant & " name=" & Session("XstemaSid") & "-" & Session("XstemaStesq") & "Y"
		sxlinkant = sxlinkant & "-" & Session("XstemaStest")
		'sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face=corflaticon>&#231;</font></a>"
		sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
		botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

		' sn=request.querystring("sn")

		If  stnivel <> 1 Then BOTONES=""



		sHTML = sHTML & "<td  rowspan=2 >&nbsp;</td>"
		sHTML = sHTML & "<td colspan=20 rowspan=2 >" & sxlink & Session("XstemaSitem") & "</a></td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>&nbsp;</td>"
		sHTML = sHTML & "<td>&nbsp;</td>"
		sHTML = sHTML & "</tr>"


		If stnivel > 1 Then
			For xxin = 2 To stnivel
				sesn = "X" & xxin
				sxlink = ""
				sxlink = sxlink & "<A class=nt HREF=Objetos.asp?Tipo=ESQUEMAS&ID="
				sxlink = sxlink & Session(sesn & "Sid")
				sxlink = sxlink & "&cantregmover=" &  Session(sesn & "Stestmover") & "&sitem="
				sxlink = sxlink & OFV.convertircar(Session(sesn & "Sitem"), " ", "_") & "&stitem="
				sxlink = sxlink & OFV.convertircar(Session(sesn & "Stitem"), " ", "_") & "&sticono="
				sxlink = sxlink & Session(sesn & "Sticono") & "&stnivel=" & Session(sesn & "Stnivel")
				sxlink = sxlink & "&stesq=" & Session(sesn & "Stesq") & "&stestado=" & Session(sesn & "Stest")
				sxlink = sxlink & " ID=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				sxlink = sxlink & "-" & Session(sesn & "Stest")
				sxlink = sxlink & " name=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				sxlink = sxlink & "-" & Session(sesn & "Stest")
				sxlink = sxlink & " TARGET=fesquemas >"


				sHTML = sHTML & "<tr>"
				sHTML = sHTML & "<td colspan=" & ((xxin - 1) * 2) & "  rowspan=2>&nbsp;</td>"

				For xind = 1 To xxin
					If xind = xxin Then
						If BrowserName() = "IE" Or BrowserName() = "CHROME" Then

							smas = "<font color=#993333 size=2><b>-</b></font>"
							sHTML = sHTML & "<td    >&nbsp;</td>"
							sHTML = sHTML & "<td style='" & borderl & borderb & "'>&nbsp;</td>"

						Else
							sHTML = sHTML & "<td width=2% >&nbsp;</td><td width=2% >"

							sHTML = sHTML & "<img ID=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq")
							sHTML = sHTML & "N"
							sHTML = sHTML & "-" & Session(sesn & "Stest")
							sHTML = sHTML & " name=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq")
							sHTML = sHTML & "N"
							sHTML = sHTML & "-" & Session(sesn & "Stest")
							sHTML = sHTML & " border=0 src=Iconos\ICorAbre.jpg>"

							sHTML = sHTML & "</td>"
						End If
					Else
					End If
				Next

				'  WMATRMEMP    = "S"
				'   WMATRMEMPACT = "S"

				If Session(sesn & "Stest") = "3" Then
					xsf = sfont
					xsff = sffont
				ElseIf Session(sesn & "Stest") = "1" And Session("Menu") = "ParaAprobar" Then
					xsf = sfont2
					xsff = sffont2
				ElseIf Session(sesn & "Stest") = "2" And Session("Menu") = "Desarrollo" Then
					xsf = sfont2
					xsff = sffont2
				Else
					xsf = ""
					xsff = ""

				End If

				If WMATRMEMP = "S" And WMATRMEMPACT = "N" Then
					xsf = sfont3
					xsff = sffont3
				End If


				snid = Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				snid = snid & "-" & Session(sesn & "Stest")



				sxiconv = "<font id='" & snid & "' face='CorFlaticon' size=3 color='#f7aa22'>&#052;</font>"


				sHTML = sHTML & "<td class=ttt colspan=2  rowspan=2 style='" & Session("CBRDICO") & "'>"

				sHTML = sHTML & sxlink
				sHTML = sHTML & sxiconv
				sHTML = sHTML & "</a></td>"

				If xsf <> "" Then xsf = replace(xsf,"XXXX",snid)


				'    sHTML = sHTML & sxlink & "<img border=0 src=Iconos\ICorAbierto.jpg"
				'    sHTML = sHTML & " ID=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				'    sHTML = sHTML & "-" & Session(sesn & "Stest")
				'    sHTML = sHTML & " name=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				'    sHTML = sHTML & "-" & Session(sesn & "Stest")
				'    sHTML = sHTML & " >"

				sHTML = sHTML & "</a></td>"
				sHTML = sHTML & "<td  rowspan=2>&nbsp;</td>"
				sHTML = sHTML & "<td colspan=" & (20 - ((xxin - 1) * 2)) & "  rowspan=2>" & sxlink & xsf
				sxlinksig =  "<A id=lnksig class=nt style='display:none;'  HREF=Objetos.asp?xbt=" & xbt & "&Tipo=ESQUEMAS&ID="
				sxlinksig = sxlinksig & Session(sesn & "Sid") & "&sitem="
				sxlinksig = sxlinksig & OFV.convertircar(Session(sesn & "Sitem"), " ", "_")
				sxlinksig = sxlinksig & "&CantRegMover=" & Session(sesn & "Stestmover") + CantregMostrar & "&stitem="
				sxlinksig = sxlinksig & OFV.convertircar(Session(sesn & "Stitem"), " ", "_") & "&sticono="
				sxlinksig = sxlinksig & Session(sesn & "Sticono") & "&stnivel=" & Session(sesn & "Stnivel")
				sxlinksig = sxlinksig & "&stesq=" & Session(sesn & "Stesq") & "&stestado=" & Session(sesn & "Stest")
				sxlinksig = sxlinksig & " ID=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				sxlinksig = sxlinksig & "-" & Session(sesn & "Stest")
				sxlinksig = sxlinksig & " name=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				sxlinksig = sxlinksig & "-" & Session(sesn & "Stest")
				'sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face=corflaticon>&#232;</font></a>"
				sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"
				' sxlink = sxlink & " TARGET=fesquemas >"





				sxlinkant =  "<A id=lnkant class=nt style='display:none;' HREF=Objetos.asp?xbt=" & xbt & "&Tipo=ESQUEMAS&ID="
				sxlinkant = sxlinkant & Session(sesn & "Sid") & "&sitem="
				sxlinkant = sxlinkant & OFV.convertircar(Session(sesn & "Sitem"), " ", "_")
				If  CantRegMover <= 0 Then
					sxlinkant = sxlinkant & "&CantRegMover=0"
				Else
					sxlinkant = sxlinkant & "&CantRegMover="  &   Session(sesn & "Stestmover")- CantregMostrar

				End If
				sxlinkant = sxlinkant   & "&stitem="
				sxlinkant = sxlinkant & OFV.convertircar(Session(sesn & "Stitem"), " ", "_") & "&sticono="
				sxlinkant = sxlinkant & Session(sesn & "Sticono") & "&stnivel=" & Session(sesn & "Stnivel")
				sxlinkant = sxlinkant & "&stesq=" & Session(sesn & "Stesq") & "&stestado=" & Session(sesn & "Stest")
				sxlinkant = sxlinkant & " ID=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				sxlinkant = sxlinkant & "-" & Session(sesn & "Stest")
				sxlinkant = sxlinkant & " name=" & Session(sesn & "Sid") & "-" & Session(sesn & "Stesq") & "N"
				sxlinkant = sxlinkant & "-" & Session(sesn & "Stest")
				'sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face=corflaticon>&#231;</font></a>"
				sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
				botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig
				'aca botones
				' CantRegMover=0
				If cstr(xxin&"")=cstr(stnivel&"") Then
					sHTML = sHTML & Session(sesn & "Sitem") & xsff & "</a>"&xbotones&"</td>"
				Else
					sHTML = sHTML & Session(sesn & "Sitem") & xsff & "</a></td>"
				End If
				sHTML = sHTML & "</tr>"

				sHTML = sHTML & "<tr>"
				sHTML = sHTML & "<td >&nbsp;</td>"
				sHTML = sHTML & "<td >&nbsp;</td>"
				sHTML = sHTML & "</tr>"

			Next
		End If

		'    StrSql = "SELECT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado as estcod, "
		'    StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.estado FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
		'    StrSql = StrSql & " where  "
		'    StrSql = StrSql & " (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') and  "
		'    StrSql = StrSql & "RELAOBJ.codojasoc=OBJETOS.codobjeto and "
		'    StrSql = StrSql & "OBJPRTY.estado=Estados.objid and "
		'    StrSql = StrSql & "OBJETOS.codobjeto = OBJPRTY.codobjeto and "
		'    StrSql = StrSql & "OBJETOS.OBJUN = " & SESSION("UN") & " and "
		'    StrSql = StrSql & " RELAOBJ.codobjmaster=" & sID
		'    StrSql = StrSql & " and relaobj.tipobjmaster = '" & stesq & "' "
		'    StrSql = StrSql & " And OBJETOS.objnivel >= " & Session("SegNivel")
		'    StrSql = StrSql & " And (OBJPRTY.estado=" & Session("sEstadon")
		'    StrSql = StrSql & " or OBJPRTY.estado=3 "
		'    If Session("Menu") = "ParaAprobar" Then
		'       StrSql = StrSql & " or OBJPRTY.estado=1 or OBJPRTY.estado=5 or OBJPRTY.estado=6 "
		'    ElseIf Session("Menu") = "Desarrollo" Then
		'       StrSql = StrSql & " or OBJPRTY.estado=2 or OBJPRTY.estado=5 or OBJPRTY.estado=6 "
		'    End If
		'    StrSql = StrSql & " ) "
		'    StrSql = StrSql & " order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, OBJETOS.codobjeto "



		StrSql = ""

		If stesq= "PEDIDO" Or stesq= "SNN" Then

			StrSql= " SELECT pedidos.objid,pedidosdet.nombre,RELAOBJ.tipobjasoc,EstadosForm.estado as estcod,Relaobj.tipobjasoc,"
			StrSql = StrSql & " '','','',pedidos.solicitante,EstadosForm.estado as estcod,pedidos.estado as estadoP,clienteplan.cliente "
			StrSql = StrSql & " FROM RELAOBJ  "

			'*** INCIDENTE 457: Se agrega join entre pedidos y relaobj por el tipo de objeto ***'
			StrSql = StrSql & " INNER JOIN pedidos  ON (RELAOBJ.codojasoc = pedidos.objid AND RELAOBJ.tipoBjasoc = pedidos.tipobjeto AND (pedidos.estado =" & Session("sEstadon") '&  " ) "

			StrSql = StrSql & " or pedidos.estado=3 "
			If Session("Menu") = "ParaAprobar" Then
				StrSql = StrSql & " or pedidos.estado=1 or pedidos.estado=5 or pedidos.estado=6 "
			ElseIf Session("Menu") = "Desarrollo" Then
				If stesq= "SNN" Then
					StrSql = StrSql & " or pedidos.estado=2 or pedidos.estado=5 or pedidos.estado=6 or pedidos.estado=4 or pedidos.estado=7"
				ElseIf stesq= "PEDIDO" Then
					StrSql = StrSql & " or pedidos.estado=2 or pedidos.estado=5 or pedidos.estado=6 "
				End If
			End If
			StrSql = StrSql & ") ) "
			StrSql = StrSql & "INNER JOIN pedidosdet  ON (pedidosdet.idpedido = pedidos.objid) "
             StrSql = StrSql & " left JOIN clienteplan  ON (clienteplan.id=pedidos.objun ) "

			If stesq = "PEDIDO" Then
				StrSql = StrSql & " INNER JOIN EstadosForm ON (pedidos.estado=EstadosForm.objid)"

			ElseIf stesq = "SNN" Then
				StrSql = StrSql & " INNER JOIN EstadosForm ON (pedidos.estado=EstadosForm.objid)"

			End If

			StrSql = StrSql & " where "
			strSql = StrSql & " relaobj.tipobjmaster = '" & stesq & "'  and RELAOBJ.codobjmaster=" & sID


			'INCIDENTE 466: Orden por objid para los pedidos de formularios, antes lo hacia por fecha de inicio. Si habia varios pedidos para la misma fecha los ordenaba de cualquier forma y no siguiendo el objid

			'CODIGO ANTERIOR
			'strSql = StrSql & " order by pedidos.fechaini desc "
			'CODIGO NUEVO

			If stesq = "PEDIDO" Then
				strSql = StrSql & " and EstadosForm.tipo='PD' "	'Se agrega filtro por tipo de pedido'
				strSql = StrSql & " order by pedidos.objid desc "

			ElseIf stesq = "SNN" Then
				strSql = StrSql & " and EstadosForm.tipo='SN' " 'Se agrega filtro por tipo de solicitud'
				strSql = StrSql & " order by pedidos.objid desc "
			End If

			'response.write  strSql
		Else

           ' StrSql =   "  SELECT    OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado as estcod, Relaobj.tipobjasoc, OBJPRTY.estado, b.codobjeto, d.codobjetol "
			'NUEVO 02/08/2021 GP --> ALE
            '10317-0004/20
			StrSql =   "  SELECT   distinct OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado as estcod, Relaobj.tipobjasoc, OBJPRTY.estado, b.codobjeto, d.codobjetol,objetos.indobj "
			'------------
			StrSql = StrSql & " FROM RELAOBJ "
			'*** INCIDENTE 457: Se agrega join entre objetos y relaobj por el tipo de objeto ***'
			StrSql = StrSql & " INNER JOIN OBJETOS ON ( RELAOBJ.codojasoc = OBJETOS.codobjeto AND RELAOBJ.tipoBjasoc = objetos.tipobjeto )"
			StrSql = StrSql & " INNER JOIN TIPOBJETO ON (TIPOBJETO.TIPOBJETO = OBJETOS.TIPOBJETO) "
			StrSql = StrSql & " INNER JOIN OBJPRTY ON (OBJETOS.codobjeto = OBJPRTY.codobjeto ) "

			'--------------------------------------------------------------------------
            'NUEVO 02/08/2021 GP --> ALE
            '10317-0004/20
			SELECT CASE session("APLMDCLI")

		    CASE "GP"

			   StrSql = StrSql & " LEFT JOIN OBJMEMPALC ON (OBJMEMPALC.codobjeto = OBJPRTY.codobjeto AND OBJMEMPALC.TIPOBJETO=OBJPRTY.TIPOBJETO) "

			END SELECT
           '--------------------------------------------------------------------------

			StrSql = StrSql & " INNER JOIN ESTADOS ON (OBJPRTY.estado=Estados.objid ) "
			'    strsql = strsql & " left join (select distinct codobjeto from OBJOBSCMT where actpend is null ) b on (OBJETOS.codobjeto = b.codobjeto) "
			StrSql = StrSql & " left join (select distinct codobjeto from OBJOBSCMT ) b on (OBJETOS.codobjeto = b.codobjeto) "
			strsql = strsql & " left join (select distinct codobjetol from OBJlnk where ltipo = 'LRGVN' ) d on (OBJETOS.codobjeto = d.codobjetol) "

			If SESSION("BASE")="K" Then
				StrSql = StrSql & " where  TIPOBJETO.TIPOCPT ='K' AND "
			Else
				StrSql = StrSql & " where  TIPOBJETO.TIPOCPT <>'K' AND "
			End If


			StrSql = StrSql & " (OBJETOS.descri like '%" & sbuscar & "%') and "
			StrSql = StrSql & " (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') and  "


           '--------------------------------------------------------------------------
            'NUEVO 02/08/2021 GP --> ALE
            '10317-0004/20
            'TRAE DOC EMPRESA 1  Y OBJMEMPALC.CODEMPRE --> DEL USUARIO
            ' O OBJETOS.OBJUN -> DEL USUARIO

			SELECT CASE session("APLMDCLI")

			CASE "GP"

			   StrSql = StrSql & " ((OBJETOS.OBJUN = " & SESSION("UN") & ") OR ( OBJETOS.OBJUN=1 AND OBJMEMPALC.CODEMPRE= " & SESSION("UN") & " ))  and "
			 '  StrSql = StrSql & " ((OBJETOS.OBJUN = " & SESSION("UN") & ") OR ( OBJETOS.OBJUN=1 AND OBJMEMPALC.CODEMPRE= " & SESSION("UN") & " ))  and "

			CASE ELSE

			   StrSql = StrSql & " OBJETOS.OBJUN = " & SESSION("UN") & " and "

			END SELECT
			'--------------------------------------------------------------------------

			StrSql = StrSql & " RELAOBJ.codobjmaster=" & sID
			StrSql = StrSql & " and relaobj.tipobjmaster = '" & stesq & "' "
			StrSql = StrSql & " And OBJETOS.objnivel >= " & Session("SegNivel")
			StrSql = StrSql & " And (OBJPRTY.estado=" & Session("sEstadon")
			StrSql = StrSql & " or OBJPRTY.estado=3 "
			If Session("Menu") = "ParaAprobar" Then
				StrSql = StrSql & " or OBJPRTY.estado=1 or OBJPRTY.estado=5 or OBJPRTY.estado=6 "
			ElseIf Session("Menu") = "Desarrollo" Then
				StrSql = StrSql & " or OBJPRTY.estado=2 or OBJPRTY.estado=5 or OBJPRTY.estado=6 "
			End If
			StrSql = StrSql & " ) "
			StrSql = StrSql & " order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, OBJETOS.codobjeto "



		End If
		sNuevoTipo = "ESQUEMAS"
		nNivel = stnivel + 1
		Session("sID2") = sID
		Session("EsqSup") = "N"
	End Select

	'response.write strsql
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")

	If rs.EOF Then
		If sID <> "xxx" Then
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td colspan=" & ((nNivel - 1) * 2) & " >&nbsp;</td>"
			For xind = 1 To nNivel
			Next
			sHTML = sHTML & "<td colspan=5 >&nbsp;</td><td class=vti colspan=" & (20 - ((nNivel - 1) * 2)) & ">"
			sHTML = sHTML & "No hay Esquemas asociados</td></tr>"
		Else
			sIcono = "ICorRecycle.jpg"
			stitem = "RYS"
			sitem = "CARPETA DE OBJETOS "
			stestado = "N"
			xsf = ""
			xsff = ""
			' ---------------------------------

			sxlink = ""
			sxlink = sxlink & "<A class=nt ID=0-RYS" & Session("EsqSup")
			sxlink = sxlink & "-" & stestado
			sxlink = sxlink & " name="
			sxlink = sxlink & "0-RYS" & Session("EsqSup")
			sxlink = sxlink & "-" & stestado
			sxlink = sxlink & " HREF=Objetos.asp?Tipo="
			sxlink = sxlink & sNuevoTipo & "&ID=0" & "&sitem="
			sxlink = sxlink & OFV.convertircar(sitem, " ", "_") & "&stitem="
			sxlink = sxlink & OFV.convertircar(stitem, " ", "_") & "&sticono=" & sIcono & "&stnivel="
			sxlink = sxlink & nNivel & "&stesq=" & "RYS" & "&stestado=" & stestado & " TARGET=fesquemas>"

			sHTML = sHTML & "<tr>"


			StrSql = "Select relaobj.relac from relaobj, OBJPRTY, Estados "
			StrSql = StrSql & "where relaobj.codojasoc = OBJPRTY.codobjeto and "
			StrSql = StrSql & "OBJPRTY.estado = Estados.objid and "
			StrSql = StrSql & " relaobj.tipobjmaster = '" & "RYS" & "'"
			StrSql = StrSql & " and relaobj.codobjmaster = 0"
			StrSql = StrSql & " And (OBJPRTY.estado=" & Session("sEstadon")
			StrSql = StrSql & " or OBJPRTY.estado=3 ) "

			Set rsmas = OFV.crearconsultaEx(StrSql, cn, 1, "")
			If Not rsmas.EOF Then
				smas = "<font color=#993333 size=1><b>+</b></font>"
			Else
				smas = "<font color=#993333 size=2><b>-</b></font>"
			End If
			Call OFV.cerrarconsulta(rsmas)

			sHTML = sHTML & "<td colspan=" & ((nNivel - 1) * 2) & "   rowspan=2>&nbsp;</td>"

			For xind = 1 To nNivel
				If xind = nNivel Then
					If BrowserName() = "IE" Or BrowserName() = "CHROME" Then

						sHTML = sHTML & "<td    >&nbsp;</td>"
						sHTML = sHTML & "<td style='" & borderl & borderb & "'>&nbsp;</td>"
					Else
						sHTML = sHTML & "<td width=2% align=right>" & smas & "</td><td width=2% >"
						sHTML = sHTML & "<img ID=0-RYS" & Session("EsqSup")
						sHTML = sHTML & "-" & stestado
						sHTML = sHTML & " name="
						sHTML = sHTML & "0-RYS" & Session("EsqSup")
						sHTML = sHTML & "-" & stestado
						sHTML = sHTML & " src=Iconos\ICorAbre.jpg"
						sHTML = sHTML & " border=0></td>"
					End If
				Else
				End If
			Next

			snid = "0-RYS" & Session("EsqSup")
			snid = snid & "-" & stestado



			sxiconv = "<font id='" & snid & "' face='CorFlaticon' size=3 color='#f7aa22'>&#052;</font>"


			sHTML = sHTML & "<td class=ttt colspan=2  rowspan=2 style='" & Session("CBRDICO") & "'>"

			sHTML = sHTML & sxlink
			sHTML = sHTML & sxiconv
			sHTML = sHTML & "</a></td>"


			'  sHTML = sHTML & sxlink & "<img ID=0-RYS"
			'  sHTML = sHTML & Session("EsqSup")
			'  sHTML = sHTML & "-" & stestado
			'  sHTML = sHTML & " name=0-RYS" & Session("EsqSup")
			'  sHTML = sHTML & "-" & stestado
			'  sHTML = sHTML & " border=0 src=Iconos\" & sIcono & "></a></td>"

			If xsf <> "" Then xsf = replace(xsf,"XXXX",snid)

			sHTML = sHTML & "<td   rowspan=2>&nbsp;</td>"
			sHTML = sHTML & "<td colspan=" & (20 - ((nNivel - 1) * 2)) & "   rowspan=2>" & sxlink & xsf & sitem
			sHTML = sHTML & xsff & "</a></td>"
			sHTML = sHTML & "</tr>"


			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td >&nbsp;</td>"
			sHTML = sHTML & "<td >&nbsp;</td>"
			sHTML = sHTML & "</tr>"

			' ---------------------------------
		End If
	Else

		xTrColor = "#F0F0F0"

		posi=1
		' response.write CantregMover
		rs.move CantregMover

		' if  rs.EOF then posi=CantregMostrar+1
		Do Until rs.EOF Or posi > CantregMostrar
			'  do while not rs.EOF and posi < CantregMostrar



			If UCase(sTipo) = "ESQUEMAS" Then

				If nv = 1 Then
					sHTML = sHTML & "<tr>"

					sHTML = sHTML & "<td >&nbsp;</td>"
					sHTML = sHTML & "<td style='" & borderl & "'>&nbsp;</td>"
					sHTML = sHTML & "</tr>"
				Else
					nv = 1
				End If

				subtprv = "N"
				If stesq="PEDIDO" Or stesq="SNN" Then
					' response.write rs(4)
					stitem = rs(4)
					stestado = rs("estadoP")
				Else
					stitem = rs(4)
					stestado = rs(5)
				End If

				If CStr(stestado) = "3" Then
					xsf = sfont
					xsff = sffont
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"
				ElseIf CStr(stestado) = "1" And Session("Menu") = "ParaAprobar" Then
					xsf = sfont2
					xsff = sffont2
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"
				ElseIf CStr(stestado) = "6" And Session("Menu") = "ParaAprobar" Then
					xsf = "<b>" & sfont6
					xsff = sffont6 & "</b>"
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"
				ElseIf CStr(stestado) = "6" And Session("Menu") = "Desarrollo" Then
					xsf = "<b>" & sfont6
					xsff = sffont6 & "</b>"
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"
				ElseIf CStr(stestado) = "2" And Session("Menu") = "Desarrollo" Then
					xsf = sfont2
					xsff = sffont2
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"
				ElseIf  CStr(stestado) = "5" And Session("Menu") = "ParaAprobar" Then
					xsf = sfont5
					xsff = sffont5
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"

				ElseIf  CStr(stestado) = "5" And Session("Menu") = "Desarrollo" Then
					xsf = sfont5
					xsff = sffont5
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"

				ElseIf  CStr(stestado) = "4" And stitem = "SN" Then
					xsf = sfont5
					xsff = sffont5
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"

				ElseIf  CStr(stestado) = "1" And stitem = "SN" Then
					xsf = sfont2
					xsff = sffont2
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"

					'***********

					'INCIDENTE 466: Se agrega Elseif para que tome el estado 1 de los pedidos que no los estaba tomando en cuenta
					'Para que se puedan ver en el arbol de ABM Formularios

				ElseIf  CStr(stestado) = "1" And stitem = "PDCORXY" Then
					xsf = sfont2
					xsff = sffont2
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"

					'*****************

					'INCIDENTE 457: Se agrega Elseif para que tome el estado 7 de las SN que no los estaba tomando en cuenta
					'Para que se puedan ver en el arbol de SN

				ElseIf  CStr(stestado) = "7" And stitem = "SN" Then
					xsf = sfont2
					xsff = sffont2
					sVige = "&nbsp;&nbsp;(" & rs(3) & ")"
				Else

					xsf = ""
					xsff = ""
					sVige = ""
				End If


				spendobs = ""

				If (rs(6) & "") <> "" And stestado <> "3" Then

					spendobs = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='Corflaticon' color='#bb0000' size=3 title='Hay observaciones pendientes.' style='background-color:#f0f0f0;cursor:pointer;' >&#093;</font>"

				End If

				spendprg = ""

				If (rs(7) & "") <> "" Then

					spendprg = "&nbsp;&nbsp;<font face='Corflaticon' color='#f7aa22' size=3 title='documento vinculado a un subproceso.' style='background-color:#f0f0f0;cursor:pointer;' >&#0110;</font>"

				End If


				If WMATRMEMP = "S" And WMATRMEMPACT = "N" Then
					xsf = sfont3
					xsff = sffont3
				End If

				sNIVE = "D"

				StrSql = "Select archivo, descri, TIPOCPT From Tipobjeto "
				StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs(4) & "'"
				Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not rsIcono.EOF Then
					sIcono = rsIcono(0)
					sTipoE = rsIcono(1)
					sNIVE  = rsIcono(2)

					'**************************

					'INCIDENTE 466: Se cambia la visualización de PD y SN. Ahora es: PD o SN, Nro, Nombre, Autor, Estado

					'***CODIGO ANTERIOR
					'IF RS(1) & "" = "" or rs(2) = "SN" THEN -- Se comenta ya que pregunta si el nombre es vacio o si el tipo es SN. Ahora directamente pregunta por el tipo SN
					'                     XX=GETINFOUSER(RS("SOLICITANTE"))
					'sitem= " Solicitud Nro: " & rs(0) &  " De "  & SESSION("SGU_NOMBRE")
					'ELSE
					'IF sNIVE="PD" or sNIVE="SN" THEN
					'sitem =   "Pedido Nro: " & left(rs(1) ,100) & "..."
					'ELSE

					'sitem = sTipoE & ": " & rs(1)
					'END IF
					'END IF

					'**************CODIGO NUEVO ********************
					'***										****
					'*** VISUALIZACION DE NOMBRES DE ESQUEMAS   ****
					'***										****
					'***********************************************

					'******************************************************
					'****   VISUALIZACION DE ABM SOLICITUDES NORMATIVAS ***
					'******************************************************

					If sNIVE="SN" Then 'Se toma la variable Snive para que sea la misma variable que el IF de pedidos
						XX=GETINFOUSER(RS("SOLICITANTE")) 'Llama a la funcion Getinfouser con el campo solicitante del query RS
						If len(rs(1)) >= 30 Then	'Se pone un límite de 30 caracteres para el nombre de la solicitud. Si es mayor o igual a 30
							sitem= " SN Nro: " & rs(0) & " - " & left(rs(1),30) & "... ["  & SESSION("SGU_NOMBRE") & "] " 'Tipo + Nro + Nombre (30) + Autor
						Else 'Sino
							sitem= " SN Nro: " & rs(0) & " - " & rs(1) & " ["  & SESSION("SGU_NOMBRE") & "] " 'Tipo + Nro + Nombre + Auto
						End If
                        if Session("EMPMUN")  = "S" then sitem =sitem  & " - " & rs("cliente")
					Else

						'********************************************************
						'****   VISUALIZACION DE ABM FORMULARIOS			  ***
						'****   VISUALIZACION DE NORMAS Y FORMULARIOS		  ***
						'********************************************************


						'IF sNIVE="PD" or sNIVE="SN" THEN
						If sNIVE="PD" Then 'Se toma la variable Snive
							XX=GETINFOUSER(RS("SOLICITANTE")) 'Llama a la funcion Getinfouser con el campo solicitante del query RS
							If len(rs(1)) >= 30 Then 'Se pone un límite de 30 caracteres para el nombre del pedido. Si es mayor o igual a 30
								sitem =   "PD Nro: " & rs(0) & " - " & left(rs(1),30) & "... [" & SESSION("SGU_NOMBRE") & "] " 'Tipo + Nro + Nombre (30) + Autor
							Else 'Sino
								sitem =   "PD Nro: " & rs(0) & " - " & rs(1) & " [" & SESSION("SGU_NOMBRE") & "] " 'Tipo + Nro + Nombre + Autor
							End If
							if Session("EMPMUN")  = "S" then sitem =sitem  & " - " & rs("cliente")

						Else 'Sino
							sitem = sTipoE & ": " & rs(1) 'Este ELSE se usa para todo el resto de los esquemas que no son pedidos... IMPORTANTE!!! 'Tipo de Esquema + Nombre
							'sitem = EsqnomPxex (rs(2),rs(0),ofv,sVige)
							sitem = EsqnomPx (rs(2),rs(0),ofv)
							sitem = sitem & sVige
						End If
					End If


					'If sNIVE = "D" Or sNIVE = "K" Or sNIVE = "PD" Or sNIVE = "SN" Then sitem = sitem & sVige
					 if  sNIVE = "PD" Or sNIVE = "SN" Then sitem = sitem & sVige

				End If

				Call OFV.cerrarconsulta(rsIcono)

			End If

			'***********************************************************

			If sNIVE = "C" Or sNIVE = "K"  Then


				StrSql = "Select * From OBJSDK Where tipobjeto='" & rs(4) & "' and codobjeto = " & rs(0)
				subtprv = "N"
				Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
				If Not rsIcono.EOF Then


					If Session("Usuario") <> rsIcono(7) And rsIcono(5) Then


						Call OFV.cerrarconsulta(rsIcono)


						StrSql = "Select * From OBJSDKDET Where tipobjeto='" & rs(4)
						StrSql = StrSql & "' and codobjeto = " & rs(0)
						StrSql = StrSql & " and ("
						StrSql = StrSql & "(grpusr = 'USR-" & Session("Usuario") & "') "

						'''''''''''''''''''''''''''''''''''''''''''
						If Session("DELLOG") = "S" Then

							XMFTGRP = Split(Session("SegGrpLst"), Chr(9))
							xcount = UBound(XMFTGRP) + 1
							XMFSEP = ""
							StrSql = StrSql & " or ( grpusr in ("

							For XMFI = 0 To UBound(XMFTGRP)
								StrSql = StrSql & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
								XMFSEP = ","
							Next

							StrSql = StrSql & ")) "

						Else
							StrSql = StrSql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')"
						End If


						''''''''''''''''''''''''''''''''''''''''''


						StrSql = StrSql & ") "

						'response.write strsql & "<br>"
						Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
						If  rsIcono.EOF Then subtprv = "S"

					End If
				End If
				Call OFV.cerrarconsulta(rsIcono)



			End If

			If UCase(sTipo) = "SUBTEMAS" Then

				wpstipo = rs(7)


				If rs(5) <> "0" Then
					subtprv = "S"
					StrSql = "Select * From subtemaskr "
					StrSql = StrSql & "Where codprodserv = " & rs(0)

					StrSql = StrSql & " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "


					'''''''''''''''''''''''''''''''''''''''''''
					If  Session("DELLOG") = "S" Then

						XMFTGRP = split(session("SegGrpLst"),chr(9))
						xcount  = ubound(XMFTGRP) + 1
						XMFSEP  = ""
						StrSql = StrSql & " or grpusr in ("

						For XMFI = 0 To ubound(XMFTGRP)
							StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
							XMFSEP  = ","
						Next

						StrSql = StrSql & ") "

					Else
						StrSql = StrSql & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
					End If


					''''''''''''''''''''''''''''''''''''''''''
					StrSql = StrSql & ")"


					Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
					If Not rsIcono.EOF Then
						subtprv = "N"
					End If
					Call OFV.cerrarconsulta(rsIcono)
				Else
					subtprv = "N"
				End If

				If subtprv <> "S" Then

					Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs(0) & " AND ATRMTIPOST = 'S' AND ATRMCODES = " & session("UN")
					Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
					If Not rsIcono.EOF Then
						If Not rsIcono(4) Then  subtprv = "S"
					End If
					Call OFV.cerrarconsulta(rsIcono)

				End If


				If subtprv <> "S" Then

					If nv = 1 Then
						sHTML = sHTML & "<tr>"

						sHTML = sHTML & "<td >&nbsp;</td>"
						sHTML = sHTML & "<td style='" & borderl & "'>&nbsp;</td>"
						sHTML = sHTML & "</tr>"
					Else
						nv = 1
					End If

					sIcono = rs(4)
					stitem = rs(3)
					sitem = rs(3) & ": " & rs(1)
					stestado = "N"
					' xsf = ""
					'xsff = ""

					xsf = sfont4
					xsff = sffont4

				End If

			End If




			If subtprv <> "S" Then

				If wpstipo = "R" Then

					sxlink = ""
					sxlink = sxlink & "<A class=nt ID='redir' "
					sxlink = sxlink & " name='redir' "
					sxlink = sxlink & " HREF='javascript:void(0);' "
					sxlink = sxlink & " TARGET=fesquemas >"

					snid = ""



				Else

					sxlink = ""
					sxlink = sxlink & "<A class=nt ID=" & rs(0) & "-" & rs(2) & Session("EsqSup")
					sxlink = sxlink & "-" & stestado
					sxlink = sxlink & " name="
					sxlink = sxlink & rs(0) & "-" & rs(2) & Session("EsqSup")
					sxlink = sxlink & "-" & stestado
					sxlink = sxlink & " HREF=Objetos.asp?Tipo="
					sxlink = sxlink & sNuevoTipo & "&ID=" & rs(0) & "&sitem="
					sxlink = sxlink & OFV.convertircar(sitem, " ", "_")
					'ver cantregmover
					sxlink = sxlink & "&CantRegMover=0"

					sxlink = sxlink &  "&stitem="
					sxlink = sxlink & OFV.convertircar(stitem, " ", "_") & "&sticono=" & sIcono & "&stnivel="
					sxlink = sxlink & nNivel & "&stesq=" & rs(2) & "&stestado=" & stestado & " TARGET=fesquemas>"
					'aca subcarpeta


					snid = rs(0) & "-" & rs(2) & Session("EsqSup")
					snid = snid  & "-" & stestado



				End If




				sHTML = sHTML & "<tr>"


				StrSql = "Select relaobj.relac from relaobj, OBJPRTY, Estados, OBJETOS   , TIPOBJETO "
				StrSql = StrSql & "where relaobj.codojasoc = OBJPRTY.codobjeto and "
				StrSql = StrSql & " TIPOBJETO.TIPOBJETO=OBJPRTY.TIPOBJETO AND "
				If SESSION("BASE")="K" Then
					StrSql = StrSql & " TIPOBJETO.TIPOCPT ='K' AND "
				Else
					StrSql = StrSql & " TIPOBJETO.TIPOCPT <> 'K' AND "
				End If
				StrSql = StrSql & " OBJPRTY.estado = Estados.objid and "
				StrSql = StrSql & " OBJPRTY.codobjeto = OBJETOS.codobjeto and "
				StrSql = StrSql & " relaobj.tipobjmaster = '" & rs(2) & "'"
				StrSql = StrSql & " and relaobj.codobjmaster = " & rs(0)
				StrSql = StrSql & " And (OBJPRTY.estado=" & Session("sEstadon")
				StrSql = StrSql & " or OBJPRTY.estado=3 "
				StrSql = StrSql & " AND (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "
				If Session("Menu") = "ParaAprobar" Then
					StrSql = StrSql & " or OBJPRTY.estado=1 or OBJPRTY.estado=5 or OBJPRTY.estado=6 "
				ElseIf Session("Menu") = "Desarrollo" Then
					StrSql = StrSql & " or OBJPRTY.estado=2 or OBJPRTY.estado=6 "
				End If
				StrSql = StrSql & " ) "

				'  RESPONSE.WRITE STRSQL

				Set rsmas = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not rsmas.EOF Then
					smas = "<font color=#993333 size=1> (+)</font>" & spendobs & spendprg
				Else
					smas = "" & spendobs & spendprg
				End If
				Call OFV.cerrarconsulta(rsmas)

				If ((nNivel - 1) * 2) > 0 Then
					sHTML = sHTML & "<td colspan=" & ((nNivel - 1) * 2) & "  rowspan=2 >&nbsp;</td>"
				End If

				For xind = 1 To nNivel
					If xind = nNivel Then

						sHTML = sHTML & "<td  >&nbsp;</td>"
						sHTML = sHTML & "<td style='" & borderl & borderb & "'>&nbsp;</td>"

						''          If OFV.explorador = "xxMSIE" Then
						''             sHTML = sHTML & "<td  >&nbsp;</td>"
						''             sHTML = sHTML & "<td style='" & borderl & borderb & "'>&nbsp;</td>"
						''          Else
						''            sHTML = sHTML & "<td width=2% style='cursor:pointer;' >"
						''            sHTML = sHTML & "<img ID=" & rs(0) & "-" & rs(2) & Session("EsqSup")
						''            sHTML = sHTML & "-" & stestado
						''            sHTML = sHTML & " name="
						''            sHTML = sHTML & rs(0) & "-" & rs(2) & Session("EsqSup")
						''            sHTML = sHTML & "-" & stestado
						''            sHTML = sHTML & " src=Iconos\ICorAbre.jpg"
						''            sHTML = sHTML & " border=0  onclick=""Editar(this.id);"" ></td><td style='" & borderl & borderb & "' >&nbsp;</td>"
						''          End If

					Else
					End If
				Next



				If  sNIVE = "D" Or sNIVE = "K" Or snive="PD" Or snive="SN" Then

					sxiconv = "<font id='" & snid & "' face='CorFlaticon' size=4 color='#993333'>&#0165;</font>"
				Else

					If wpstipo = "R" Then

						sxiconv = "<font id='" & snid & "' face='CorFlaticon' size=4 color='#99ccee'>&#0204;</font>"

					Else

						sxiconv = "<font id='" & snid & "' face='CorFlaticon' size=4 color='#336699'>&#0204;</font>"

					End If

				End If





				''   sHTML = sHTML & "<td class=ttt colspan=2  rowspan=2 style='" & Session("CBRDICO") & "'>"
				''   sHTML = sHTML & sxlink
				''   sHTML = sHTML & sxiconv
				''   sHTML = sHTML & "</a></td>"


				sHTML = sHTML & "<td id='" & snid & "' class=uts colspan=2  rowspan=2 style='" & Session("CBRDICO") & ";cursor:pointer;' onclick=""Editarx(this.id);"" >"
				sHTML = sHTML & "<a id='" & snid & "' class=nt href='javascript:void(0);' >" & sxiconv   &"</a>"
				sHTML = sHTML & "</td>"


				If xsf <> "" Then xsf = replace(xsf,"XXXX",snid)

				If sNIVE <> "D" And sNIVE <> "K" And sNIVE <> "PD" And sNIVE <> "SN" Then

					sitemfl = "<u >" & xsf & sitem & xsff & "</u>"  & smas
				Else

					sitemfl = xsf & sitem & smas & xsff ' &snid

				End If



				sHTML = sHTML & "<td   rowspan=2>&nbsp;</td>"
				sHTML = sHTML & "<td bgcolor=" & xTrColor & "  colspan=" & (20 - ((nNivel - 1) * 2)) & "   rowspan=2  bgcolor=" & xTrColor & "  >" & sxlink
				sHTML = sHTML & sitemfl
				sHTML = sHTML & "</a></td>"
				sHTML = sHTML & "</tr>"

				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F0F0F0"
				Else
					xTrColor = "#FFFFFF"
				End If


			End If

			rs.movenext
			posi=posi+1
		Loop
	End If

	' Call OFV.cerrarconsulta(rs)
	'Call OFV.cerrarconn(cn)

	If nv = 1 Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td >&nbsp;</td>"
		sHTML = sHTML & "<td >&nbsp;</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=1% >"
	sHTML = sHTML & "<td width=76% >"
	sHTML = sHTML & "<td width=1% >"

	sHTML = sHTML & "<tr>"

	sHTML = sHTML & "</table>" ' & botones
	If  Not rs.eof Or CantregMover > 0 Then
		'  sHTML = sHTML & ofv.botoneraESSX("2",rs,5,10,"no","no",10,"","","no","",""," ")
		sHTML = sHTML & "<table  ><tr><td>&nbsp;</td><td><font size=2 color=red face=tahoma>Ver Más</font>" & botones  & tablab
		sHTML = sHTML & "</td></tr></table>"
	End If
	If Not rs.eof Then
		sHTML = sHTML & "<script>document.getElementById('lnksig').style.display='';</script>"
		'sHTML = sHTML & "<script>document.getElementById('lnkant').style.display='';</script>"

	End If

	If CantregMover > 0 Then
		sHTML = sHTML & "<script>document.getElementById('lnkant').style.display='';</script>"
	End If
	sHTML = sHTML & "</body></html>"
	Generar = sHTML





	'''''''''''

	sHTML = Generar
	Call OFV.cerrarconsulta(rs)
	Call OFV.cerrarconn(cn)





End If

response.write sHTML


%>
