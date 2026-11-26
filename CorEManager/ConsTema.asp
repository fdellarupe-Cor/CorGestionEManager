<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires = 0
    sli=""
	session("ModMain") = "CL"
	session("xorigen") = "tema"

	XXXX = NEWEMPRE("NO")
	'RESPONSE.WRITE "2"
	If Session("CorApli") = "COREM20" Then
		wkaplic = "EM"
	ElseIf Session("CorApli") = "COREP20" Then
		wkaplic = "PL"
	Else
		wkaplic = "OR"
	End If

	sEMPRE        = request.querystring("EMPRE")
	If  SEMPRE    = "" Then
		SEMPRE    = ATREMP(0,0)
	Else
		recarga   = NEWEMPRE(SEMPRE)
		'RESPONSE.WRITE "22222"
	End If

	sTipo         = request.querystring("Tipo")
	sUN           = SEMPRE

	If sUN = "" Then sUN = 1
	session("UN") = sUN


	response.expires=-1441
	Response.AddHeader "PRAGMA", "NO-CACHE"
	sTIPOT=request.querystring("TIPOT")


	LNombre = request.querystring("Nombre")
	sNombre = request.querystring("Nombre")
	sNombre = ofv.convertircar(sNombre,"_"," ")

	sID     = request.querystring("ID")

	sitem   = ofv.convertircar(request.querystring("sitem"),"_"," ")
	stitem  = ofv.convertircar(request.querystring("stitem"),"_"," ")
	sticono = request.querystring("sticono")
	stnivel = request.querystring("stnivel")
	TIPOVISTA=request.querystring("tipovista")
	sformato=request.querystring("FORMATO")

	If request.querystring("Menu") <> "" Then Session("Menu") = request.querystring("Menu")
	If sNombre = "" Then sNombre = UCase(sTipo)

	ssTipo                = sTipo
	session("sNivelCero") = sTipo
	session("sEstado")    = "APROBADA"
	session("sEstadon")   = "3"
	If sTIPOT & ""  <> "" Then
		session("tipot")=sTIPOT
	ElseIf  sformato="ABR" Then
		sTIPOT=session("tipot")
	End If

	'response.write request.querystring


	'response.write request.querystring & "<br>" & request.form
	Set cn = ofv.conectar(ofv.strconn2)
	Strsql = "Select estado from estados where objid = " & session("sEstadon")
' response.write strsql & "<br>"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		ssTipo = rs(0)
	End If
	Call ofv.cerrarconsulta(rs)

	t="for (var i = 0; i < window.top.frames.length; i++) {"
	t=t & "		     alert(window.frames[i]);"
	t=t & "}"

	If sformato="ABR" Then
		WSPDIVN = "DTTEMAINT"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
		sHTML = ""
		sHTML = sHTML & "<HTML>"
		' sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"
		sHTML = sHTML & "<BODY onload=""window.top.document.getElementById('DTTEMAINT').innerHTML=window.top.document.getElementById('xxxx').contentWindow.document.body.innerHTML"" >"

	Else
		sHTML      = ofv.FormHeader(session("FormN"))
	End If

	shtml=shtml & "<script language=javascript>" & Chr(13)

	shtml=shtml & "function chequear(){" & Chr(13)
	shtml=shtml & "alert(window.top.CorType + window.top.name);" & Chr(13)
	shtml=shtml & "}" & Chr(13)

	shtml=shtml & "</script>" & Chr(13)

	sImagenI   = "<B><FONT FACE=""CorFlaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
	sImagenI2  = "<B><FONT FACE=""CorFlaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
	sImagenI3  = "<B><FONT FACE=""CorFlaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	sImagenEM  = "<B><FONT FACE=CorFlaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
	sImagen4si = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>4</FONT></B>"
	sImagenx   = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"


	sNombre = ofv.convertircar(sNombre," ","_")

	sHTML   = sHTML '& "<center>"

	sstipo  = session("FormN")
	HrfAsp  = "ObjTreeSubTemaCons.asp?formato=N&"
	'  HrfAsp  = "Constemaold.asp?"
	HrfTrg  = ""

	If session("empmun") = "S" Then

		Fila    = ""

		SINPUT3 = "&nbsp;<A class=tt title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

		Fila    = ""
		Celda   = ""
		sHref   = ""

		sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
		Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
		Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
		Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

		Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

		Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")

		shtml=shtml & "<script language=javascript>" & Chr(13)




		shtml=shtml & "function shwDem(){" & Chr(13)
		shtml=shtml & "document.getElementById('DEM1').style.display='';" & Chr(13)
		shtml=shtml & "}" & Chr(13)
		shtml=shtml & "</script>" & Chr(13)
		'----------------------------------------------
		sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "document.getElementById('DEM1').style.display='none';" & Chr(34)
		sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"

		'''  TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:100px;width:250px;height:250px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"

		'''  TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")

		TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:220px;left:20px;height:280px;width:380px;border-style:outset;border-color:#888888;border-width:2;border-radius: 15px 15px 15px 15px;overflow:auto;'>"

		TABLAX = TablaX & OFV.gentabla("", "aaa style='border-radius: 15px 15px 15px 15px;' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")

		TABLAX = TablaX & "<tr><td  "
		TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
		TABLAX = TablaX & "<tr><td colspan=3 "
		TABLAX = TablaX & " >&nbsp;</td></tr>"





		For IAEMP=1 To TAEMP

			'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

			If ATREMP(IAEMP,2) = 1 Then

				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "<tr  >"
				Else
					TABLAX = TablaX & "<tr>"
				End If

				TABLAX = TablaX & "<td width=10% "
				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					'   TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
					xbord = " style='border-style:solid;border-width:1;' "
					ximg = sImagen4si & "&nbsp;"
				Else
					'    TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
					xbord = "  "
					ximg = sImagenI3 & "&nbsp;"
				End If




				TABLAX = TablaX & "<td class=vti width=60% " & xbord & " >"
				'''    sHref = "OBJtree.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
				sHref = "ConsTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0) &"&TIPOVISTA="
				sLink = ofv.GenLink("","cc",sHref,"","","",ximg & "<b>" & Replace(ATREMP(IAEMP,4) & ""," ","&nbsp;") & "</b>","si")
				TABLAX = TablaX & sLink
				TABLAX = TablaX & "</td>"

				sImage = "<img src='../COREss20/Empresa/" & ATREMP(IAEMP,5) & "' width='100' >"
				'''  sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
				TABLAX = TablaX & "<td width=40% "
				TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

				TABLAX = TablaX & "</tr>"
			End If

		Next

		TAEMP=0

		For IAEMP=1 To TAEMP

			If ATREMP(IAEMP,2) = 1 Then

				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
				Else
					TABLAX = TablaX & "<tr>"
				End If

				TABLAX = TablaX & "<td width=5 "
				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
				Else
					TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
				End If

				sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
				TABLAX = TablaX & "<td width=5 "
				TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

				TABLAX = TablaX & "<td class=vt  >"

				sHref = "ConsTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0) &"&TIPOVISTA=" &"&tipot=" & sTIPOT
				sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
				TABLAX = TablaX & sLink
				TABLAX = TablaX & "</td>"
				TABLAX = TablaX & "</tr>"
			End If

		Next


		TABLAX = TablaX & OFV.genFtabla()

		TABLA = Tabla & TABLAX

		TABLA = Tabla & "</DIV>"

	End If



	TABLA = Tabla & "<br>"
	TABLA = Tabla & "<DIV id='x2' STYLE='display:BLOCK;z-index:1;margin-top:1;margin-left:1;overflow:auto;'>"
	Fila = ""




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
	If sTIPOT="PEDIDO" Then
		Strsql = "Select DISTINCT A.* From Tema  A "
		STRSQL =STRSQL & " INNER JOIN PRODSERV B ON(A.CODTEMA=B.CODTEMA AND B.MARCAPS='PEDIDO')"
		STRSQL =STRSQL & " where A.PSTIPO ='K' AND (A.Aplic = '" & wkaplic & "' or A.aplic = 'GLB') "
		Strsql = STRSQL & " order by A.inda, A.descri"
	Else
        StrSqlX  = " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

			 '''''''''''''''''''''''''''''''''''''''''''
			             if  Session("DELLOG") = "S" then

			                 XMFTGRP = split(session("SegGrpLst"),chr(9))
			                 xcount  = ubound(XMFTGRP) + 1
			                 XMFSEP  = ""
			                 StrSqlx = StrSqlx & " or grpusr in ("

			                 for XMFI = 0 to ubound(XMFTGRP)
			                     StrSqlX = StrSqlX  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
			                     XMFSEP  = ","
			                 next

			                 StrSqlX = StrSqlX & ") "

			             else
			                 StrSqlX = StrSqlX & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
			             end if


			 ''''''''''''''''''''''''''''''''''''''''''

              StrSqlX  = StrSqlX & ")"

		If session("base")="K" Then
			Strsql = "Select distinct a.* From Tema a "
			Strsql = strsql & " inner join prodserv b on (a.codtema = b.codtema and b.marcaps <> 'PEDIDO' and b.marcaps <> 'SNN') "
			Strsql = strsql & " where a.PSTIPO = 'K' AND (a.aplic = '" & wkaplic & "' or a.aplic = 'GLB') order by a.inda, a.descri"


			Strsql ="Select   DISTINCT D.codtema,D.descri,D.pstipo,D.psurl,D.aplic,D.inda "
			Strsql = strsql & " From "
			Strsql = strsql & " Prodserv A"
			Strsql = strsql & " INNER JOIN SUBTEMA B ON (A.marcaps = B.tipobjeto   )"
			Strsql = strsql & " INNER JOIN TEMA D ON (D.CODTEMA= A.CODTEMA )"
			Strsql = strsql & "  LEFT JOIN subtemaskr C ON (A.codprodserv = C.codprodserv "& StrSqlX &")"

			Strsql = strsql & " where   a.marcaps <> 'PEDIDO' and a.marcaps <> 'SNN' and (A.SEGACESO <> '1' OR (C.OBJID IS NOT NULL)) and d.PSTIPO = 'K' AND (d.aplic = '" & wkaplic & "' or d.aplic = 'GLB')"
            Strsql = strsql & " order by d.inda, d.descri"



		Else
			'Strsql = "Select distinct a.* From Tema a "
			'Strsql = strsql & " inner join prodserv b on (a.codtema = b.codtema and b.marcaps <> 'PEDIDO' and b.marcaps <> 'SNN') "
			'Strsql = strsql & " where a.PSTIPO <> 'K' AND (a.aplic = '" & wkaplic & "' or a.aplic = 'GLB') order by a.inda, a.descri"


			 Strsql ="Select   DISTINCT D.codtema,D.descri,D.pstipo,D.psurl,D.aplic,D.inda "
			 Strsql = strsql & " From "
			 Strsql = strsql & "Prodserv A"
			 Strsql = strsql & " INNER JOIN SUBTEMA B ON (A.marcaps = B.tipobjeto  )"
			 Strsql = strsql & " INNER JOIN TEMA D ON (D.CODTEMA= A.CODTEMA )"
			 Strsql = strsql & "  LEFT JOIN subtemaskr C ON (A.codprodserv = C.codprodserv "& StrSqlX &")"

			Strsql = strsql & " where  a.marcaps <> 'PEDIDO' and a.marcaps <> 'SNN'  and (A.SEGACESO <> '1' OR (C.OBJID IS NOT NULL)) and d.PSTIPO <> 'K' AND (d.aplic = '" & wkaplic & "' or d.aplic = 'GLB')"
            Strsql = strsql & " order by d.inda, d.descri"

		End If

	End If
	'response.write Strsql & "<br>"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If  rs.eof Then

		Celda   = ""
		Celda   = Celda & ofv.GenCelda("","ttt","","2%","","","","&nbsp;","si")
		Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","","No hay elementos asociados","si")

		Fila    = fila & ofv.GenRow("","","","","30","","",Celda,"si")

	Else

		Celda   = ""


		sLink   = "<button value='ver' onclick='chequear();' >"
		Celda   = Celda & ofv.GenCelda("","ttt colspan=2 ","","2%","","","",sLink,"si")
		'     Fila    = ofv.GenRow("","aaa ","","","20","","",Celda,"si")



		Do Until rs.eof



			WTVER = 1

			Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs(0) & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
			' response.write strsql & "<br>"
			Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  Not rs1.eof Then
				If Not RS1(4) Then  WTVER = 0
				'  response.write strsql & " " & WTVER & "<br>"
			End If

			If  WTVER = 1 Then

				Celda   = ""

				If rs(3) = "R" Then

					sHref   = rs(4) & " "

					If InStr(UCase(sHref),"HTTP://") = 0 Then sHref   = "HTTP://" & sHref



					'sHref   = HrfAsp & "Tipo=" & sNuevoTipo & "&ID=" & rs(0) & "&sitem=TEMA:_"
					'sHref   = sHref & ofv.convertircar(rs(1)," ","_") & "&stitem="
					'sHref   = sHref & ofv.convertircar(stitem," ","_") & "&sticono=" & sIcono
					'sHref   = sHref & "&stnivel=" & nNivel
					'sImage  = "<font face='corflaticon' size=3 color='#990000'>4</font>"
					'           sLink   = ofv.GenLink("","nt",sHref,"_blank","","",sImage,"si")
					'    Celda   = Celda & ofv.GenCelda("","ttt ","center","2%","","","",sLink,"si")

					If rs(5) = "GLB" Then
						sImage = "<font face='CorFlaticon' size=2 color='#f7aa22'>&#0252;</font>"
					Else
						sImage = "<font face='CorFlaticon' size=2 color='#f7aa22'>&#060;</font>"
					End If
					sImage  = "<font face='corflaticon' size=3 color='#990000'>4</font>"
					'     sLink   = ofv.GenLink("","nt",sHref,"_self","","",sImage,"si")
					'  Celda   = Celda & ofv.GenCelda("","ut  ","center","2%","","","",sLink,"si")

					'  sImage  = "<font face='corflaticon' size=3 color='#990000'>4</font>"
					'  sLink   = ofv.GenLink("","nt",sHref,"_blank","","",sImage,"si")

					Celda   = Celda & ofv.GenCelda("","ut ","center","2%","","","",sLink,"si")

					sLink   = ofv.GenLink("","nt",sHref,"_blank","","",rs(1),"si")
					Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","90%","","","",sLink,"si")

					' sLink   = ofv.GenLink("","nt style='color:#000000;' ",sHref,"_self","","",Capitalize(rs(1)),"si")
					'   sLink = ofv.GenLink("","nt",sHref,"_self","","",rs(1),"si")
					'    Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","90%","","","",sLink,"si")


				Else

					sHref   = HrfAsp & "Tipo=" & sNuevoTipo & "&ID=" & rs(0) & "&sitem=TEMA:_"
					sHref   = sHref & ofv.convertircar(rs(1)," ","_") & "&stitem="
					sHref   = sHref & ofv.convertircar(stitem," ","_") & "&sticono=" & sIcono
					sHref   = sHref & "&stnivel=" & nNivel &"&tipovista=" & tipovista



					If rs(5) = "GLB" Then
						sImage = "<font face='CorFlaticon' size=2 color='#f7aa22'>&#0252;</font>"
					Else
						sImage = "<font face='CorFlaticon' size=2 color='#f7aa22'>&#060;</font>"
					End If

					If sformato="ABR" Then
						sLink   = ofv.GenLink("","nt",sHref,"xxxx","","",sImage,"si")
					Else
						sLink   = ofv.GenLink("","nt",sHref,"_self","","",sImage,"si")
					End If


					' sLink   = ofv.GenLink("","nt",sHref,"_self","","",sImage,"si")
					Celda   = Celda & ofv.GenCelda("","vt  ","center","2%","","","",sLink,"si")

					'sLink   = ofv.GenLink("","nt style='color:#000000;font-size:8pt;' ",sHref,"_self","","",Capitalize(rs(1)),"si")
					'  sLink = ofv.GenLink("","nt",sHref,"_self","","",rs(1),"si")
					If sformato="ABR" Then
						sLink = ofv.GenLink("","nt",sHref,"xxxx","","",rs(1) ,"si")
					Else
						sLink = ofv.GenLink("","nt",sHref,"_self","","",rs(1) ,"si")
					End If

					Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","90%","","","",sLink,"si")

				End If
                'smasq = "<div style='float:left'><i class='las la-folder-plus la'></i></div>"
                'slink="<div class='Dvindice' >"  & sLink & smas & "</div>"
                sli= sli &  chr(13) & armaCont(slink)  'DocFormat("headerrrr",sLink,"","headerw","contenido","footer","tema","")' "<div class='Dvindice' >"  & sLink & "</div>"
               ' DocFormat("header",scontenido,sfooter,hclass,cclass,fclass,tipo,xtmas)

				Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")

				If  xTrColor = "#FFFFFF" Then
					xTrColor = "#F9F9F9"
				Else
					xTrColor = "#FFFFFF"
				End If

				nv      = 1

			End If

			rs.movenext
		Loop

	End If


	TABLA = Tabla & ofv.GenTabla("","aaa  style='border-style=solid;border-width:1;border-color:#eeeeee;' ","","90% ","","0","","-2","4","","",Fila,"si")

	TABLA = Tabla & "</DIV>"

	Call ofv.cerrarconsulta(rs)


	'Botonera
	Agregar = ""
	Buscar  = ""
	Celda   = ""

	If session("SMPLMENU") = "S" Then
		sAccion = "novedmenu.asp"
		'     Celda   = ofv.BotonVolver(sAccion)

		xxxx    = CorGenBotonXp(16, "NO", "home", "tahoma", "NO", "Home", "_self", "Home")
		Celda  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")


	End If


	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

    xcontenedor= "<div class='contenedorcons'  >"

	xcontenedor=xcontenedor &     "<div class='headertema'>TEMAS</div>"
    xcontenedor=xcontenedor &     "<div class='contenido'  >"
    xcontenedor=xcontenedor & "<div width=100% >  "


    xcontenedor=xcontenedor &  sli   & "</div>"

   xcontenedor=xcontenedor &     "</div>"
  xcontenedor=xcontenedor &"  </div>"

	sHTML = sHTML & xcontenedor & "</body></html>"
End If
Call ofv.cerrarconn(cn)
response.write sHTML


Function INDICE()
	session("sEstado")    = "APROBADA"
	session("sEstadon")   = "3"

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



	sVerEstado2 = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

	If SESSION("EMPMUN") = "S" Then

		If SESSION("GLBMUN") = 1 Then
			sverestado2 = sverestado2 & " and Objetos.OBJUN = " & Session("UN") & " "
		Else
			sverestado2 = sverestado2 & " and "
			sverestado2 = sverestado2 & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
			sverestado2 = sverestado2 & "   ( OBJETOS.OBJUN = 1 AND "
			sverestado2 = sverestado2 & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
			sverestado2 = sverestado2 & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
			sverestado2 = sverestado2 & " ) "
		End If

	End If


	sverestado2 = sverestado2 & " and Objprty.Estado = 3 "

	xStrSqla = ""
	xStrSqla  = xStrSqla & " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

	'''''''''''''''''''''''''''''''''''''''''''
	If  Session("DELLOG") = "S" Then

		XMFTGRP = split(session("SegGrpLst"),chr(9))
		xcount  = ubound(XMFTGRP) + 1
		XMFSEP  = ""
		xStrSqla = xStrSqla & " or grpusr in ("

		For XMFI = 0 To ubound(XMFTGRP)
			xStrSqla = xStrSqla  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
			XMFSEP  = ","
		Next

		xStrSqla = xStrSqla & ") "

	Else
		xStrSqla = xStrSqla & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
	End If

	''''''''''''''''''''''''''''''''''''''''''
	xStrSqla  = xStrSqla & ")"







	If SESSION("APLMDCLI") = "BGAL" Then


		StrSqlX  =  " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

		'''''''''''''''''''''''''''''''''''''''''''
		If  Session("DELLOG") = "S" Then

			XMFTGRP = split(session("SegGrpLst"),chr(9))
			xcount  = ubound(XMFTGRP) + 1
			XMFSEP  = ""
			StrSqlx = StrSqlx & " or grpusr in ("

			For XMFI = 0 To ubound(XMFTGRP)
				StrSqlX = StrSqlX  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
				XMFSEP  = ","
			Next

			StrSqlX = StrSqlX & ") "

		Else
			StrSqlX = StrSqlX & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
		End If


		''''''''''''''''''''''''''''''''''''''''''

		StrSqlX  = StrSqlX & ")"

		Strsql = "Select distinct  a.codtema,a.descri,a.pstipo,a.psurl,a.aplic, a.inda   From Tema a "
		Strsql = strsql & " inner join prodserv b on (a.codtema = b.codtema and b.marcaps <> 'PEDIDO' and b.marcaps <> 'SNN') "
		Strsql = strsql & " where a.PSTIPO <> 'K' AND (a.aplic = '" & wkaplic & "' or a.aplic = 'GLB') "
		Strsql = strsql & " union "
		Strsql = strsql & " Select distinct a.codtema,a.descri,a.pstipo,a.psurl,a.aplic, a.inda From Tema a "
		Strsql = strsql & " where (a.PSTIPO <> 'K' and a.PSTIPO = 'R') "
		Strsql = strsql & " order by a.inda, a.descri "
		'response.write strsql


		Strsql = " Select DISTINCT D.codtema,D.descri,D.pstipo,D.psurl,D.aplic,D.inda"
		Strsql = strsql & "  From "
		Strsql = strsql & " Prodserv A"
		Strsql = strsql & " INNER JOIN SUBTEMA B ON (A.marcaps = B.tipobjeto  )"
		Strsql = strsql & " INNER JOIN TEMA D ON (D.CODTEMA= A.CODTEMA )"
		Strsql = strsql & " LEFT JOIN subtemaskr C ON (A.codprodserv = C.codprodserv  " & StrSqlX &")"

		Strsql = strsql & " where   a.marcaps <> 'PEDIDO' and a.marcaps <> 'SNN' and a.PSTIPO <> 'K' AND (d.aplic = '" & wkaplic & "' or d.aplic = 'GLB')  and (A.SEGACESO <> '1' OR (C.OBJID IS NOT NULL))"



		Strsql = strsql & " union "
		Strsql = strsql & " Select distinct D.codtema,D.descri,D.pstipo,D.psurl,D.aplic, D.inda From Tema d "
		Strsql = strsql & " where (D.PSTIPO <> 'K' and D.PSTIPO = 'R') "
		Strsql = strsql & " ORDER BY  D.inda, D.descri"
		' response.write strsql


	Else

		Strsql = "Select codtema,descri,pstipo,psurl,aplic, inda  From Tema "
		Strsql = strsql & " where (aplic = '" & wkaplic & "' or aplic = 'GLB') and PSTIPO IN ('C','R') order by inda, descri"

	End If


	' response.write strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If  rs.eof Or session("un") = 0 Then

		Celda   = ""
		Celda   = Celda & ofv.GenCelda("","ttt","","2%","","","","&nbsp;","si")
		Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","","No hay elementos asociados","si")

		Fila    = fila & ofv.GenRow("","","","","25","","",Celda,"si")

	Else

		Celda   = ""



		Fila=""
		codtema=0
		isubtema=0
		ISUBCARPETA=""

		If (SESSION("UN") & "") = "" Then


			xxxx = NEWEMPRE("NO")


			SEMPRE = ATREMP(0, 0)
			SESSION("UN") = SEMPRE & ""

			If (SESSION("UN") & "") = "" Then SESSION("UN") = 1

		Else
			'  xxxx = NEWEMPRE(SESSION("UN"))


		End If

		cant=1
		sul=""
		sli=""

		Do until rs.eof

			xxnova = "S"
			Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs("codtema") & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
			'      response.write strsql
			Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			If Not  rs1.eof Then
				If rs1(4) = 0 Then xxnova = "N"
			End If

			If xxnova = "S" Then
				sxlink = "ObjTreeSubTemaCons.asp?formato=N&"
				sHref  = ""
				'   if codtema <> rs("codtema") then
				'   codtema=rs("codtema")

				If rs("aplic") = "GLB" Then
					'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#0252;&</font>&nbsp;&nbsp;"
				Else
					'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#060;</font>&nbsp;&nbsp;"
					'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'></font>&nbsp;&nbsp;"
				End If


				'    xTrColor="#eeeeff"
				sitem2   = ofv.ConvertirCar(rs("DESCRI")," ","_")
				sxlink  = sxlink & "ID=" & rs("codtema") & "&sitem=" & sitem2 & "&stitem=" & REPLACE(stitem," ","_")
				sxlink  = sxlink & "&stnivel=0&tipo=SUBTEMAS&TIPOVISTA=CONS"
				sHref   = sxlink

				starget=""
				If rs("pstipo")="R" Then
					'sImage  = "<font face='corflaticon' size=3 color='#990000'>4</font>"
					'sImage  = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#990000'>4</font>&nbsp;&nbsp;"
					sHref   = rs("psURL") & " "
					If instr(ucase(sHref),"HTTP://") = 0 Then sHref   = "HTTP://" & sHref
					starget=" targer=_new"
				End If


				sLink2   = ofv.GenLink("","'nt xnti'" ,sHref,starget,"","",replace(rs("DESCRI")," "," ") & " ","si")


				'celda= ofv.GenCelda("","'ut uti' xbgcolor='" & xTrColor & "' xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","",sImage & slink2,"si")

				'Fila    = Fila & ofv.GenRow("","'aaa rowI' ","","","20","","",Celda,"si")
				 sli= sli & armaCont(slink2) '& "<div class='Dvindice' >"  & sLink2 & "</div>"




			End If
			RS.MOVENEXT
		Loop



			indice=  sli

	End If






	Call OFV.cerrarconsulta(rs)





	'INDICE2 = ofv.GenTabla("","aaa   ","","100% ","","0","","0","0","","",Fila,"si")
End Function

function armaCont(dato)


stdato=""

stdato="<div class='idxv-item'>"
stdato=stdato & "  <div class='content' style='flex:1;min-width:0;'>"
stdato=stdato & "    <div class='title-row'>"
stdato=stdato & "        <span class='type-chip'>"
stdato=stdato & "          <i class='la la-folder'></i>"
stdato=stdato & "          <span class='chip-name'>" & dato & "</span>"
stdato=stdato & "        </span>"
'stdato=stdato & "        <button class='child-inline' title='Abrir subnivel'><i class='la la-plus'></i></button>"
stdato=stdato & "      </span>"
stdato=stdato & "      <div class='right-ctrl'>"
'stdato=stdato & "        <span class='badge'>Vigente</span>"
'stdato=stdato & " 		<button class='icon-btn info' title='Ver descripción'><i class='la la-info-circle'></i></button>"
'stdato=stdato & "    	<button class='icon-btn fav active' title='Favorito'><i class='la la-star'></i></button>"
'stdato=stdato & "    	<button class='icon-btn sub active' title='Suscripto'><i class='la la-bell'></i></button>"
stdato=stdato & "        <div class='inline-actions'>"
stdato=stdato & "   	<button class='act' title='Crear'><i class='la la-plus'></i></button>"
stdato=stdato & "        <button class='act' title='Editar'><i class='la la-pen'></i></button>"
stdato=stdato & "        <button class='act' title='Eliminar'><i class='la la-trash'></i></button>"
stdato=stdato & "        <button class='act' title='Copiar'><i class='la la-copy'></i></button>"
stdato=stdato & "        <button class='act' title='Mover'><i class='la la-exchange-alt'></i></button>"
stdato=stdato & "        </div>"
stdato=stdato & "      </div>"
stdato=stdato & "    </div>"
stdato=stdato & "  </div>"

stdato=stdato & " </div>"
armaCont=stdato
end function
%>