<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

Response.expires = 0
session("SETIDEMPRE")=""
Response.expires = 0
'response.write "SESSION UN: " & session("UN") & "   " & now()
SESSION("BASE")="D"

fondo = " bgcolor=" & Session("CMFMNBOT")


sHTML = "<html><body width=400 " & fondo & "  >"

XSINICIA = REQUEST.QUERYSTRING("INICIA")
If XSINICIA = "" Then XSINICIA = "N"


If session("WEBDESK") = "S" Then

	xscantmenu = 99
	styl = ""
	styl = "background-color:" & Session("CMFMENU") & ";"
	ctablas = 0
	ctabcol = 99

Else

	If  Session("FULLGUI") = "N" Then
		styl = "background-color:#888888;"
	Else
		styl = XspEx("#888888","#cccccc",0) & ";"
	End If



End If


''     XtabBG = "background-color: rgba(76, 76, 201, 0.3);"
XtabBG = ""
XcelBG = "background-color: rgba(255, 75, 00, 0.6);"


If  session("SMPLMENU") <> "S" Then

	xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"

End If

Menu = Request.querystring("Menu")
If  Menu <> "I" Then
	Submenu  = Request.querystring("ID")
	Menu     = OFV.convertircar(Menu, "_", " ")
	sMenu    = "Menu " & Menu
	Session("cormenu") = Menu

	Fila = ""

	Dim aAplicacion()
	Dim aLinks()
	Dim aLinksImg()
	Dim aLinksLet()
	Dim aLinksTLT()
	Dim aLinksURL()
	Dim aLinksTRG()
	Dim aLinksDSC()


	sHTML = sHTML & "<table  style='" & XtabBG & "' >"
	sHTML = sHTML & "<tr>"

	xlfun = " onclick=""despmenu(); "" "




	Set cx = OFV.conectar(OFV.strconn0)


	Set cn = OFV.conectar(OFV.strconn1)


	maxUL = 6
	menucnt = 0
	Celda = ""


	StrSql = "Select * From Menux Where menuatr='" & Submenu & "' and mcodapli = '" & Session("CorApli") & "' "

	'*******Nuevo

	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not rs.EOF Then
		XMenuant = rs(1)

		'''  sImagenx3  =   "&nbsp;<font color='#993333' >&nbsp;<b>[" & XMenuant & "]</b></font>&nbsp;"


		'     Celda = sLink



		'     if menucnt = 0 then
		smenu   = "<div id=menucor><nav><h2><i class='la la-bars'></i>Menu</h2><UL>"
		'   end if

		menucnt = menucnt + 1


		''''  sLink   = "<span id='" & ssMnn & "'  class='ntmnn' >" & sInput & "</span> "
		'  sInput = "<IMG border=0 width='20px' src='../coress20/iconos/alphanumeric-crossing.png' onmouseover=""document.getElementById('DTMNV').style.display='none';"" >"
		'  sLink   = "<span   class='ntmnn' >&nbsp;" & sInput & "&nbsp;<font color='#993333' >&nbsp;<b>[" & XMenuant & "]</b></font><br>&nbsp;</span> "

		'  Celda   = Celda & "<LI style='list-style-type: none;' >"
		'  Celda   = Celda & sLink
		' Celda   = Celda & "</LI  >"

		If menucnt = maxUL Then
			'     Celda   = Celda & "</UL  >"
			'      menucnt = 0
		End If

	End If
	xsel = 1
	StrSql = "Select * From Menux Where mdepen='" & Submenu & "' and mcodapli = '" & Session("CorApli") & "' "
	' response.write strsql
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not rs.EOF Then
		If rs.fields.Count > 6 Then
			xsel = 6
			Call OFV.cerrarconsulta(rs)
			StrSql = "Select * From Menux Where mdepen='" & Submenu & "' and mcodapli = '" & Session("CorApli") & "' order by Morden "

			Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
		Else
			xsel = 1
		End If
	End If


	Do Until rs.EOF



		atrok = "N"
		StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'M' "
		StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
		StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
		StrSql = StrSql & "  Aplicacion = '" & Session("CorApli") & "'"

		Set rs1 = OFV.crearconsultaEx(StrSql, cx, 1, "")
		If rs1.EOF Then
			atrok = "O"
		ElseIf rs1(1) Then
			atrok = "S"
		Else
			atrok = "N"
		End If
		Call OFV.cerrarconsulta(rs1)

		xcount  = 0

		If atrok = "O" Then

			StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'M' "
			StrSql = StrSql & " And alias = '" & rs(xsel) & "' And  "

			If  Session("DELLOG") = "S" Then

				XMFTGRP = Split(session("SegGrpLst"),Chr(9))
				xcount  = UBound(XMFTGRP) + 1
				XMFSEP  = ""
				StrSql = StrSql & " grupo in ("

				For XMFI = 0 To UBound(XMFTGRP)
					StrSql  = StrSql  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
					XMFSEP  = ","
				Next

				StrSql = StrSql & ") "

			Else
				StrSql = StrSql & " grupo = '" & Session("seggrp") & "' "
			End If


			StrSql = StrSql & " And Aplicacion = '" & Session("CorApli") & "'"
			Set rs2 = OFV.crearconsultaEx(StrSql, cx, 1, "")
			If  rs2.EOF Then
				atrok = "O"
			Else
				If  Session("DELLOG") = "S" Then
					XMFSUM = 0
					XMFSUMN = 0
					Do Until rs2.eof
						XMFSUM = XMFSUM + 1
						If rs2(1) Then
							atrok = "S"
							Exit Do
						Else
							atrok   = "N"
							XMFSUMN = XMFSUMN + 1
						End If
						rs2.movenext
					Loop

					If  atrok = "N" Then
						If  xcount >  XMFSUMN Then atrok = "S"
					End If

				Else
					If rs2(1) Then
						atrok = "S"
					Else
						atrok = "N"
					End If
				End If
			End If
			Call OFV.cerrarconsulta(rs2)

		End If


		If atrok = "O" Or atrok = "S" Then


			sInput = Trim(rs(xsel))
			sHref = Trim(rs(2))
			sTarget = "info"
			sPath = Session("CorPath")
			If rs(3) <> "" Then sTarget = rs(3)
			spt = InStr(rs(2), ".asp")
			xesm = "N"
			If spt > 1 Then
				If Left(rs(2), spt - 1) = "Submenus" Then
					sPath = "../coress20/"
					xesm = "S"
				End If
			End If
			sseg = InStr(sHref, "?")
			If sseg = 0 Then
				sHref = sHref & "?"
			Else
				sHref = sHref & "&"
			End If
			sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")
			ssMnn = "MNN" & OFV.convertircar(UCase(rs(1)), " ", "_")

			If  session("SMPLMENU") <> "S" Then


				If session("WEBDESK") = "S"  Then



					ctabcol = ctabcol + 1

					If ctabcol > xscantmenu Then

						If ctabcol < 99 Then

							xsclase = "ntmnn2 style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;color:#cc0000;cursor:pointer;' "
							xlfun = " onclick=""document.getElementById('DBSL" & (ctablas + 1) & "').style.display='block';document.getElementById('DBSL" & (ctablas) & "').style.display='none'; "" "


						End If

						ctabcol = 1
						ctablas = ctablas + 1


						If ctablas > 1 Then



							styl2 = "display:block;"

							''  styl2 = "display:none;"
							xsclase = ""' "nt style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;color:#cc0000;cursor:pointer;' "
							xlfun = " onclick=""document.getElementById('DBSL" & (ctablas - 1) & "').style.display='block';document.all.DBSL" & (ctablas) & ".style.display='none'; "" "




						Else

							styl2 = "display:block;"
							sInput  = "<font color='#ff0000' size=2 >&#0124;</font>"

						End If



					End If


					sInput  = "&nbsp;&nbsp;" & Replace(Capitalize(UCase(rs(1) & ""))," ","&nbsp;") & "&nbsp;&nbsp;"

					xsclase =""' "nt STYLE='font-size:7pt;cursor:pointer;text-decoration:none;'  "



					'    if menucnt = 0 then
					'      Celda   = Celda & "<UL style='float:left;' >"
					'  end if

					menucnt = menucnt + 1


					sLink   = "<span id='" & ssMnn & "'  class='ntmnn' >" & sInput & "</span> "

					Celda   = Celda & "<LI  >"
					Celda   = Celda & OFV.GenLink("", "ntmnn  onclick=" & Chr(34) & "SetMnn('" & ssMnn & "');" & Chr(34) & " style='cursor:pointer;'  ", sPath & sHref, sTarget, "", "", sLink,"si")
					Celda   = Celda & "</LI  >"

					'   if menucnt = maxUL then
					'     Celda   = Celda & "</UL  >"
					'    menucnt = 0
					'  end if

				Else


					If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

						sInput  = UCase(rs(1))
						xsclase = "nt STYLE='border-bottom-style:solid;border-bottom-color:#dddddd;border-bottom-width:1;padding-left:10px;' "

					Else

						sInput  = UCase(rs(1))
						xsclase = "nt STYLE='border-bottom-style:solid;border-bottom-color:#dddddd;border-bottom-width:1;padding-left:10px;' "

					End If


					sLink = xsimagenf & OFV.GenLink("", "ntmnn2 " , sPath & sHref, sTarget, "", "", sInput, "si")
					Celda = OFV.gencelda("", xsclase, "", "", "", "", "", sLink, "si")
					Fila = Fila & OFV.genrow("", "", "", "", "20", "", "", Celda, "si")



				End If



			End If


		End If




		rs.movenext
	Loop
	Call OFV.cerrarconsulta(rs)


	'  if menucnt > 0 then
	Celda   = celda &  "<li><UL style='float:left;' >"  & Celda & "</UL  ></li>"
	'   menucnt = 0
	' end if
	sHTML = sHTML & "<td valign='top' style='" & XcelBG & "border-style:solid;border-width:1;border-color:white; border-radius: 15px 15px 15px 15px; box-shadow: 2px 2px 5px grey;' >"
	smenu = smenu & celda
	'sHTML = sHTML & "</td>"






Else


	If  session("SMPLMENU") <> "S" Then
		sHTML = sHTML & "<center><br><br>"
		sHTML = sHTML & "</center>"
	End If

End If

''''''''''''''''''''''''''''''''''''''''''''' arma menu solapas ''''''''''''

' xxxxxxxxxxxxx

celda = ""

StrSql = "Select Menux.* From Menux, CorAplic "
StrSql = StrSql & " where CorAplic.codapli = Menux.mcodapli and "
StrSql = StrSql & "CorAplic.CodApli = '" & Session("CorApli") & "' AND Menux.mdepen = '0'"
Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
If Not rs.EOF Then
	If rs.fields.Count > 6 Then
		xsel = 6
		Call OFV.cerrarconsulta(rs)

		StrSql = "Select Menux.* From Menux, CorAplic "
		StrSql = StrSql & " where CorAplic.codapli = Menux.mcodapli and "
		StrSql = StrSql & "CorAplic.CodApli = '" & Session("CorApli") & "' AND Menux.mdepen = '0' order by Menux.Morden"
		Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	Else
		xsel = 1
	End If
End If
Fila = ""
cnm = 0
menucnt = 0

' if menucnt = 0 then
'   Celda   = Celda & "<UL style='float:left;' >"
' end if

menucnt = menucnt + 1



sLink   = "<span   class='ntmnn' >&nbsp;<font color='#993333'>&nbsp;<b>OPCIONES</b></font><br>&nbsp;</span> "

Celda   = Celda & "<LI style='list-style-type: none;' >"
Celda   = Celda & sLink
Celda   = Celda & "</LI  >"

'   if menucnt = maxUL then
'     Celda   = Celda & "</UL  >"
'     menucnt = 0
'  end if

Do Until rs.EOF





	atrok = "N"
	StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'S' "
	StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
	StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
	StrSql = StrSql & "  Aplicacion = '" & Session("CorApli") & "'"

	Set rs1 = OFV.crearconsultaEx(StrSql, cx, 1, "")
	If rs1.EOF Then
		atrok = "O"
	ElseIf rs1(1) Then
		atrok = "S"
	Else
		atrok = "N"
	End If
	Call OFV.cerrarconsulta(rs1)

	xcount  = 0

	If atrok = "O" Then

		StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'S' "
		StrSql = StrSql & " And alias = '" & rs(xsel) & "' And  "

		If  Session("DELLOG") = "S" Then

			XMFTGRP = Split(session("SegGrpLst"),Chr(9))
			xcount  = UBound(XMFTGRP) + 1
			XMFSEP  = ""
			StrSql = StrSql & " grupo in ("

			For XMFI = 0 To UBound(XMFTGRP)
				StrSql  = StrSql  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
				XMFSEP  = ","
			Next

			StrSql = StrSql & ") "

		Else
			StrSql = StrSql & " grupo = '" & Session("seggrp") & "' "
		End If


		StrSql = StrSql & " And Aplicacion = '" & Session("CorApli") & "'"
		Set rs2 = OFV.crearconsultaEx(StrSql, cx, 1, "")
		If  rs2.EOF Then
			atrok = "O"
		Else
			If  Session("DELLOG") = "S" Then
				XMFSUM = 0
				XMFSUMN = 0
				Do Until rs2.eof
					XMFSUM = XMFSUM + 1
					If rs2(1) Then
						atrok = "S"
						Exit Do
					Else
						atrok   = "N"
						XMFSUMN = XMFSUMN + 1
					End If
					rs2.movenext
				Loop

				If  atrok = "N" Then
					If  xcount >  XMFSUMN Then atrok = "S"
				End If

			Else
				If rs2(1) Then
					atrok = "S"
				Else
					atrok = "N"
				End If
			End If
		End If
		Call OFV.cerrarconsulta(rs2)

	End If




	If atrok = "O" Or atrok = "S" Then
		If Celda = "N" Then
			Celda = ""
			sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
			Formulario = OFV.GenForm("initaplx", "aaa STYLE='display:none;' ", "../coress20/Submenus.asp?ID=" & rs(6) & "&Menu=" & Replace(UCase(rs(1)), " ", "_") & "&INICIA=" & XSINICIA, "Master", "post", "", sInput, "si")
			scp = "<script language='javascript' > document.initaplx.submit(); </script>"
		End If


		ctabcol = ctabcol + 1

		If ctabcol > xscantmenu Then

			If ctabcol < 99 Then

				If  session("SMPLMENU") <> "S" Then

					If session("WEBDESK") = "S" Then

					Else


						xsclase = "ntslp style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;color:#cc0000;cursor:pointer;' "

						xlfun = " onclick=""document.getElementById('DBML" & (ctablas + 1) & "').style.display='block';document.all.DBML" & (ctablas) & ".style.display='none'; "" "

						'  Celda = Celda & OFV.gencelda("BAT" & CTABLAS & "S*", xsclase, "center title='siguiente' " & xlfun, "1", "", "", "", "<b>4</b>", "si")

						sHTML = sHTML & OFV.genrow("", "", "", "", "20", "", "", celda, "si")
						sHTML = sHTML & OFV.GenFTabla()

					End If

				End If

				Celda = ""


			End If

			ctabcol = 1

			ctablas = ctablas + 1


			If ctablas > 1 Then

				If  session("SMPLMENU") <> "S" Then



					If session("WEBDESK") = "S" Then



					Else

						'''   styl2 = "display:none;"

						xsclase = "ntslp style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;color:#cc0000;cursor:pointer;' "

						xlfun = " onclick=""document.getElementById('DBML" & (ctablas - 1) & "').style.display='block';document.all.DBML" & (ctablas) & ".style.display='none'; "" "

						'''     Celda = OFV.gencelda("BAT" & CTABLAS & "A*", xsclase, "center title='anterior' " & xlfun, "1", "", "", "", "<b>3</b>", "si")

						styl2 = "display:block"
						sInput  = xfssepar
						Celda = OFV.gencelda("", "ut", "center", "1%", "", "", "", sInput, "si")

					End If



				End If

			Else


				If session("WEBDESK") = "S" Then



				Else


					''     styl2 = "display:block;"

					styl2 = "display:block;" '& session("CMFMENU") & ";"
					sInput  = xfssepar
					Celda = OFV.gencelda("", "ut", "center", "1%", "", "", "", sInput, "si")

				End If



			End If

			If  session("SMPLMENU") <> "S" Then

				If session("WEBDESK") = "S" Then

					'''   sHTML = sHTML & OFV.GenTabla("DBML" & ctablas, "aaa style='" & styl & styl2 & "' ", "", "100", "", "0", "", "-2", "1", "", "", "", "no")

					sHTML = sHTML & ""

				Else


					If ofv.ModTitCls = "1"  Then

						sHTML = sHTML & OFV.GenTabla("DBML" & ctablas, "aaa style='" & styl & styl2 & "' ", "", "100%", "", "0", "", "-2", "1", "", "", "", "no")

					Else

						sHTML = sHTML & OFV.GenTabla("DBML" & ctablas, "aaa style='" & styl & styl2 & "' ", "", "100%", "", "0", "", "-2", "1", "", "", "", "no")

					End If

				End If

			End If


		End If

		If  session("SMPLMENU") <> "S" Then

			xsfuncion = ""

			If session("WEBDESK") = "S" Then

				' xsfuncion = " onclick=" & chr(34) & "submenu();  " & chr(34) & " "
			End If


			If session("WEBDESK") = "S" Then

				ssMnn     = "SLP" & Replace(UCase(rs(1)), " ", "_")

				xGenConti = "<b>&nbsp;&nbsp;" & Replace(Capitalize(UCase(rs(1) & ""))," ","&nbsp;") & "&nbsp;&nbsp;</b>"
				sInput    = xGenConti
				xsclase   = "" '"vt style='font-size:8pt;cursor:pointer;text-decoration:none;' "

				sInput   = "<span id='" & ssMnn & "'  class='ntslp' >" & sInput & "</span> "
				sHref   = "../coress20/Submenus.asp?ID=" & rs(6) & "&Menu=" & Replace(UCase(rs(1)), " ", "_")
				sLink   = OFV.GenLink("", "ntslp  onclick=" & Chr(34) & "SetSlp('" & ssMnn & "');" & Chr(34) & " style='cursor:pointer;'  ", sHref, "Master", "", "", sInput, "si")


				'     if menucnt = 0 then
				'       Celda   = Celda & "<UL style='float:left;list-style-position: inside;' >"
				'     end if

				menucnt = menucnt + 1

				'''     Celda   = Celda & "<span  >" & xfssepar & "</span> "

				Celda   = Celda & "<LI  >"
				Celda   = Celda & sLink
				'"<span id='" & ssMnn & "'  class='XX' >" & sInput & "</span> "
				Celda   = Celda & "</LI  >"

				'      if menucnt = maxUL then
				'        Celda   = Celda & "</UL  >"
				'        menucnt = 0
				'     end if


			Else



				If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
					xGenConti = UCase(rs(1))
					sInput    = xGenConti
					xsclase   = "uti style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' "

				Else

					sInput    = UCase(rs(1))
					xsclase   = "uti style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' "
				End If

				sHref = "../coress20/Submenus.asp?ID=" & rs(6) & "&Menu=" & Replace(UCase(rs(1)), " ", "_")
				sLink = OFV.GenLink("", "aaa ", sHref, "Master", "", "", sInput, "si")
				Celda = Celda & OFV.gencelda("", xsclase & " " & xsfuncion , "", "", "", "", "", sLink, "si")

			End If

		End If

	End If






	rs.movenext
Loop
Call OFV.cerrarconsulta(rs)

'   if menucnt > 0 then
'     Celda   = Celda & "</UL  >"
'     menucnt = 0
'  end if

' sHTML = sHTML & "<td valign='top' style='" & XcelBG & "border-style:solid;border-width:1;border-color:white; border-radius: 15px 15px 15px 15px; box-shadow: 2px 2px 5px grey;' >"
' Celda   =  "<UL style='float:left;' >" & Celda & "</ul>"
'sHTML = sHTML & celda & "</li><li>"
' sHTML = sHTML & "</td>"




''''''''''''''''''' fin menu solapas ''''''''''''''''''''''''''



'''''''''''''''''''' app bar

Celda   = ""
menucnt = 0

StrSql = "select * from CorAplic"
Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
If Not rx.EOF Then
	If rx.fields.Count > 8 Then
		Call OFV.cerrarconsulta(rx)
		StrSql = "select * from CorAplic where enable <> 0 order by ordenapli "
		Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
		ind = 0

		' if menucnt = 0 then
		'    Celda   = Celda & "<UL style='float:left;' >"
		' end if

		menucnt = menucnt + 1



		sLink   = "<span   class='ntmnn' >&nbsp;<font color='#993333'>&nbsp;<b>APPs</b></font><br>&nbsp;</span> "

		Celda   = Celda & "<LI style='list-style-type: none;' >"
		Celda   = Celda & sLink
		Celda   = Celda & "</LI  >"

		' if menucnt = maxUL then
		'    Celda   = Celda & "</UL  >"
		'   menucnt = 0
		'end if


		Do Until rx.EOF
			ReDim Preserve aAplicacion(ind)
			ReDim Preserve aLinks(ind)
			ReDim Preserve aLinksImg(ind)
			ReDim Preserve aLinksLet(ind)
			ReDim Preserve aLinksTLT(ind)
			ReDim Preserve aLinksURL(ind)
			ReDim Preserve aLinksTRG(ind)
			ReDim Preserve aLinksDSC(ind)

			aAplicacion(ind) = rx(1)
			aLinks(ind) = "<a class=bt title='" & rx(12)
			aLinks(ind) = aLinks(ind) & "' href='../COREss20/" & rx(11)
			If UCase(rx(11)) = UCase("Menus.asp") Then
				aLinks(ind) = aLinks(ind) & "?Opcion=" & UCase(rx(1))
			End If
			aLinks(ind) = aLinks(ind) & "' target='" & rx(10)
			aLinks(ind) = aLinks(ind) & "' >"
			aLinksImg(ind) = "<img border=0 src='../COREss20/Imagenes/" & rx(6) & "' >"
			apllet = Replace(rx(2), "*", "")
			aLinksLet(ind) = "<b>" & Replace(apllet, " ", "&nbsp;") & "</b>"

			'                     aLinks(ind) = aLinks(ind) & "</a>"

			aLinksTLT(ind) = rx(12)
			aLinksURL(ind) = "../COREss20/" & rx(11)
			If UCase(rx(11)) = UCase("Menus.asp") Then
				aLinksURL(ind) = aLinksURL(ind) & "?Opcion=" & UCase(rx(1)) & "&INICIA=S"

			Else
				aLinksURL(ind) = rx(11)
			End If
			aLinksTRG(ind) = rx(10)
			aLinksDSC(ind) = UCase(Replace(rx(2), "*", ""))

			ind = ind + 1
			rx.movenext
		Loop


		''''''''''''''''''''''''''''' app bar

		xcount = 0
		scp    = ""



		For n = 0 To UBound(aAplicacion)

			atrok = "N"
			StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'A' "
			StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
			StrSql = StrSql & "  Aplicacion = '" & aAplicacion(n) & "'"
			Set rs1 = OFV.crearconsultaEx(StrSql, cn, 1, "")
			If rs1.EOF Then
				atrok = "O"
			ElseIf rs1(1) Then
				atrok = "S"
			Else
				atrok = "N"
			End If
			Call OFV.cerrarconsulta(rs1)

			If atrok = "O" Then


				StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'A' "
				StrSql = StrSql & " And  "

				If  Session("DELLOG") = "S" Then

					XMFTGRP = Split(session("SegGrpLst"),Chr(9))

					xcount = UBound(XMFTGRP) + 1
					XMFSEP  = ""
					StrSql = StrSql & " grupo in ("

					For XMFI = 0 To UBound(XMFTGRP)
						StrSql  = StrSql  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
						XMFSEP  = ","
					Next

					StrSql = StrSql & ") "

				Else
					StrSql = StrSql & " grupo = '" & Session("seggrp") & "' "
				End If


				StrSql = StrSql & " And Aplicacion = '" & aAplicacion(n) & "'"
				Set rs1 = OFV.crearconsultaEx(StrSql, cN, 1, "")
				If  rs1.EOF Then
					atrok = "O"
				Else
					If  Session("DELLOG") = "S" Then
						XMFSUM = 0
						XMFSUMN = 0
						Do Until rs1.eof
							XMFSUM = XMFSUM + 1
							If rs1(1) Then
								atrok = "S"
								Exit Do
							Else
								atrok   = "N"
								XMFSUMN = XMFSUMN + 1
							End If
							rs1.movenext
						Loop

						If  atrok = "N" Then
							If  xcount >  XMFSUMN Then atrok = "S"
						End If

					Else
						If rs1(1) Then
							atrok = "S"
						Else
							atrok = "N"
						End If
					End If
				End If
				Call OFV.cerrarconsulta(rs1)

			End If




			If atrok = "O" Or atrok = "S"  Then
				If XFondo = "N" Then
					XFondo = n
					XFdsc  = aLinksDSC(n)
				End If

				If  sopcion = aAplicacion(n) Then
					XFondo = n
					XFdsc  = aLinksDSC(n)
				End If

				sInput  = "&nbsp;&nbsp;" & Replace(Capitalize(aLinksDSC(n))," ","&nbsp;") & "&nbsp;&nbsp;"

				xsclase =""' "nt STYLE='font-size:7pt;cursor:pointer;text-decoration:none;'  "
				sTarget = "Footer"
				sPath   = "../COREss20/"
				sHref   = "footer.asp?Opcion=" & aAplicacion(n)


				If menucnt = 0 Then
					Celda   = Celda & "<UL style='float:left;' >"
				End If

				menucnt = menucnt + 1

				sLink   = "<span  class='ntmnn' >" & sInput & "</span> "

				Celda   = Celda & "<LI  >"
				Celda   = Celda & OFV.GenLink("", "ntmnn  style='cursor:pointer;'  ", sPath & sHref, sTarget, "", "", sLink,"si")
				Celda   = Celda & "</LI  >"

				If menucnt = maxUL Then
					Celda   = Celda & "</UL  >"
					menucnt = 0
				End If







			End If

		Next

		'     if menucnt > 0 then
		'        Celda   = Celda & "</UL  >"
		'        menucnt = 0
		'     end if
		' sHTML = sHTML & "<td valign='top' style='" & XcelBG & "border-style:solid;border-width:1;border-color:white; border-radius: 15px 15px 15px 15px; box-shadow: 2px 2px 5px grey;' >"
		'   Celda = "<ul>" & celda & "</ul>"
		' sHTML = sHTML & celda


		'sHTML = sHTML & ARMAM() & "</td>"


		''''''''''''''''''''''''''''''


	Else



	End If
End If
Call OFV.cerrarconsulta(rx)

''''''''''''''''''''' fin app bar

Call OFV.cerrarconn(cn)
Call OFV.cerrarconn(cx)

sHTML = sHTML & "</tr>"
sHTML = sHTML & "</table>" ' & ARMAM()



Fila = ""

xsaccion = "../coremanager/novedmenu.asp"
'xsaccion = "nulo.asp"

sInput = OFV.GenerarInput("boton", "", "hidden", "", "1", "it", "")
Formulario = OFV.GenForm("nulo", " aa style='display:none;' ", xsaccion, "info", "post", "", sInput, "si")

'  XSINICIA="N"
If session("WEBDESK") = "S" And XSINICIA = "N" Then


	'''''    Formulario = Formulario & "<script language='javascript' > window.top.LDFOBJMENU('Master','MM2') </script>"
	Formulario = Formulario & "<script language='javascript' > window.top.LDFOBJMENU('Master','TXHI3MNV') </script>"

Else

	Formulario = Formulario & "<script language='javascript' > document.nulo.submit(); </script>"

End If

sHTML = sHTML & Formulario


shtml = shtml   & "</body></html>"



Response.Write sHTML

%>
