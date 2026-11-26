<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
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
	
	
	
	xlfun = " onclick=""despmenu(); "" "
	
	If  session("SMPLMENU") <> "S" Then
		
		If session("WEBDESK") = "S"  Then
			
		Else
			
			sImagenx3  = "<B><FONT FACE=corflaticon COLOR='#ffffff' SIZE=1>2</FONT></B>"
			sImagenx3  =   "&nbsp;" & sImagenx3 & "<font color='#ffffff' >" & "&nbsp;" & Session("cormenu") & "</font>"
			Celda =  OFV.gencelda("", "at STYLE='" & styl & ";border-style:solid;border-color:#ffffff;border-width:1;padding-left:3px;' title='Ocultar solapas.' " & xlfun, "", "", "", "", "", sImagenx3, "si")
			Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")
		End If
		
	End If
	
	
	Set cx = OFV.conectar(OFV.strconn0)
	
	
	Set cn = OFV.conectar(OFV.strconn1)
	StrSql = "Select * From Menux Where menuatr='" & Submenu & "' and mcodapli = '" & Session("CorApli") & "' "
	
	'*******Nuevo
	
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not rs.EOF Then
		XMenuant = rs(1)
		
		sImagenx3  =   "&nbsp;<font color='#993333' >&nbsp;<b>[" & XMenuant & "]</b></font>&nbsp;"
		Celda =  OFV.gencelda("", "at STYLE='" & styl & ";border-style:solid;border-color:#ffffff;border-width:1;padding-left:3px;' title='Ocultar solapas.' " & xlfun, "", "", "", "", "", sImagenx3, "si")
		Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")
		
	End If
	xsel = 1
	StrSql = "Select * From Menux Where mdepen='" & Submenu & "' and mcodapli = '" & Session("CorApli") & "' "
	
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
							
							'   Celda = Celda & OFV.gencelda("BAT" & CTABLAS & "S*", xsclase, "center title='siguiente' " & xlfun, "1", "", "", "", "<b>4</b>", "si")
							
							
							' sHTML = sHTML & OFV.genrow("", "", "", "", "20", "", "", celda, "si")
							' sHTML = sHTML & OFV.GenFTabla()
							
							Celda = ""
							
						End If
						
						ctabcol = 1
						ctablas = ctablas + 1
						
						
						If ctablas > 1 Then
							
							
							
							styl2 = "display:block;"
							Celda = ""
							
							''  styl2 = "display:none;"
							xsclase = ""' "nt style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;color:#cc0000;cursor:pointer;' "
							xlfun = " onclick=""document.getElementById('DBSL" & (ctablas - 1) & "').style.display='block';document.all.DBSL" & (ctablas) & ".style.display='none'; "" "
							
							'   Celda = OFV.gencelda("BAT" & CTABLAS & "A*", xsclase, "center title='anterior' " & xlfun, "1", "", "", "", "<b>3</b>", "si")
							
							sInput  = "<font color='#ff0000' size=2 >&#0124;</font>"
							' Celda = Celda & OFV.gencelda("", "nt", "center", "1%", "", "", "", sInput, "si")
							
							
						Else
							
							styl2 = "display:block;"
							sInput  = "<font color='#ff0000' size=2 >&#0124;</font>"
							'  Celda = OFV.gencelda("", "nt", "center", "1%", "", "", "", sInput, "si")
							
						End If
						
						'  sHTML = sHTML & OFV.GenTabla("DBSL" & ctablas, "aaa style='" & styl & styl2 & "' ", "", "100", "", "0", "", "-2", "1", "", "", "", "no")
						
						
						
					End If
					
					
					sInput  = "&nbsp;&nbsp;" & Replace(Capitalize(UCase(rs(1) & ""))," ","&nbsp;") & "&nbsp;&nbsp;"
					xsclase =""' "vt STYLE='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;'  "
					
					''   sInput = OFV.gencelda("", xsclase, "center", "10%", "", "", "", sInput, "si")
					
					xsclase =""' "nt STYLE='font-size:7pt;cursor:pointer;text-decoration:none;'  "
					
					''   Celda = Celda & OFV.GenLink("", xsclase, sPath & sHref, sTarget, "", "", sInput, "si")
					'  Fila = Fila & OFV.genrow("", "", "", "", "20", "", "", Celda, "si")
					
					''   sInput  = "<font color='#ff0000' size=2 >&#0124;</font>"
					''   Celda = Celda & OFV.gencelda("", "nt", "center", "1%", "", "", "", sInput, "si")
					
					'''   sLink   = OFV.GenLink("", "cmb  onclick=" & chr(34) & "SetMnn('" & ssMnn & "');" & chr(34) & " style='font-size:8pt;cursor:pointer;text-decoration:none;color:#000000;font-weight:normal;'  ", sPath & sHref, sTarget, "", "", sInput, "si")
					
					Celda   = Celda & "<span  ><font color='#ff0000' size=2 >&#0124;</font></span> "
					sLink   = "<span id='" & ssMnn & "'  class='ntmnn' >" & sInput & "</span> "
					Celda   = Celda & OFV.GenLink("", "ntmnn  onclick=" & Chr(34) & "SetMnn('" & ssMnn & "');" & Chr(34) & " style='cursor:pointer;'  ", sPath & sHref, sTarget, "", "", sLink, "si")
					
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
	Call OFV.cerrarconn(cn)
	Call OFV.cerrarconn(cx)
	
	If session("WEBDESK") = "S"  Then
		
		sHTML = sHTML & celda &  "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"
		
		'' sHTML = sHTML & OFV.genrow("", "", "", "", "20", "", "", celda, "si")
		'' sHTML = sHTML & OFV.GenFTabla()
		
	Else
		If  session("SMPLMENU") <> "S" Then
			If ofv.ModTitCls = "1" Then
				sHTML = sHTML & OFV.GenTabla("DBSL", "aaa ", "", "100%", "", "0", "", "-2", "-2", "", "", Fila, "si")
			Else
				sHTML = sHTML & OFV.GenTabla("DBSL", "", "", "100%", "", "0", "", "-2", "-2", "", "", Fila, "si")
			End If
		End If
	End If
	
	
	
	
Else
	
	
	If  session("SMPLMENU") <> "S" Then
		sHTML = sHTML & "<center><br><br>"
		sHTML = sHTML & "</center>"
	End If
	
End If

Fila = ""

xsaccion = "../coremanager/novedmenu.asp"
'xsaccion = "nulo.asp"

sInput = OFV.GenerarInput("boton", "", "hidden", "", "1", "it", "")
Formulario = OFV.GenForm("nulo", " aa style='display:none;' ", xsaccion, "info", "post", "", sInput, "si")

'  XSINICIA="N"
If session("WEBDESK") = "S" And XSINICIA = "N" Then
	
	
	'''''    Formulario = Formulario & "<script language='javascript' > window.top.LDFOBJMENU('Master','MM2') </script>"
	Formulario = Formulario & "<script language='javascript' > window.top.LDFOBJMENU('Master','TXHI3M') </script>"
	
Else
	
	Formulario = Formulario & "<script language='javascript' > document.nulo.submit(); </script>"
	
End If

sHTML = sHTML & Formulario


shtml = shtml  & "</body></html>"



Response.Write sHTML

%>
