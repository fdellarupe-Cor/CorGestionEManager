
<%

Function GENFMTGENERAL(FNCODOBJ,FNTIPOBJ)


	FNSHTML = ""




	ReDim tobj(0)
	ReDim tobjd(0)


	oi = 0
	strsql = "SELECT * from tipobjeto"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		Do Until rs.eof
			ReDim Preserve tobj(oi)
			ReDim Preserve tobjd(oi)
			tobj(oi) = rs(0)
			tobjd(oi) = rs(1)
			oi = oi + 1
			rs.movenext
		Loop
	End If

	Call ofv.cerrarconsulta(rs)

	'xxxxxxxxxxxx
	ReDim Preserve sPath(0)
	i = 1
	Do

		strsql = "select * from csbarr where codojasoc = " & sCodigo & " and tipobjasoc = '" & sTipo & "'"
		'response.write strsql
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs.eof Then
			ReDim Preserve sPath(i)
			sCodigo = rs(1)
			sTipo = rs(2)
			sPath(i) = ""
			If Not IsNull(IsNull(rs(0))) Then
				For oi=0 To UBound(tobj)
					If rs(2) = tobj(oi) Then
						sPath(i) = tobjd(oi) & ": "
						Exit For
					End If
				Next
				sPath(i) = sPath(i) & rs(0)
			Else
				i = i + 1
				Exit Do
			End If
			i = i + 1
		Else
			Exit Do
		End If
		Call ofv.cerrarconsulta(rs)
	Loop

	strsql = "SELECT PRODSERV.descri, TEMA.descri FROM TEMA, RELAOBJ, "
	strsql = strsql & "PRODSERV where RELAOBJ.codobjmaster = PRODSERV.codprodserv and "
	strsql = strsql & " RELAOBJ.tipobjmaster = PRODSERV.marcaps and "
	strsql = strsql & "TEMA.codtema = PRODSERV.codtema and codprodserv = " & sCodigo


	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		stSubtema = rs(0)
		stTema = rs(1)
		Call ofv.cerrarconsulta(rs)
	End If

	For h = UBound(sPath) To 1 Step - 1
		If Trim(spath(h)) <> "" Then
			If h > 1 Then
				sRoot = sRoot & ofv.convertircar(sPath(h)," ","&nbsp;") & "\"
			Else
				sRoot = sRoot & ofv.convertircar(sPath(h)," ","&nbsp;")
			End If
		End If
	Next
	If Right(sRoot,1) = "\" Then sRoot = Left(sRoot,Len(sRoot)-1)

	'xxxxxxxxxx


	Strsql="Select * From Objetos Where codobjeto=" & nCodigo


	Set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

	If Not rs.eof Then
		sModifica = session("Modificar")
		vuelta= ofv.vueltaasp2("../")

		wdescri = rs(8)

		'----------------------------------------------

		sinput2 = "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=corflaticon size=1   >r</font>"

		FNSHTML = FNSHTML & "<DIV id='cx1' class='DINDM' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;"
		FNSHTML = FNSHTML & "position:absolute;left:100;top:160;width:600px;height:250px;overflow:hidden;'>"

		'''''''''''''''''''''

		FNSHTML = FNSHTML & OFV.gentabla("", "TWIND", "", "100% ", "", "0", "", "-1", "-2", "", "", "", "no")
		FNSHTML = FNSHTML & "<tr>"
		FNSHTML = FNSHTML & "<td CLASS=mbt width=99%  >&nbsp;&nbsp;ORIGEN&nbsp;DE&nbsp;LA&nbsp;DOCUMENTACION</td>"
		FNSHTML = FNSHTML & "<td CLASS=mbt width=1% ONCLICK=""document.all.cx1.style.display='none';"" title='Ocultar.' >"
		FNSHTML = FNSHTML & sinput2
		FNSHTML = FNSHTML & "</td>"
		FNSHTML = FNSHTML & "</tr>"
		FNSHTML = FNSHTML & OFV.genFtabla()
		FNSHTML = FNSHTML & "<center>"
		FNSHTML = FNSHTML & OFV.gentabla("", "", "", "98%", "", "0", "", "0", "0", "", "", "", "no")
		FNSHTML = FNSHTML & "<tr height=22 ><td width=40% "
		FNSHTML = FNSHTML & "class=ttt  >Origen</td>"
		FNSHTML = FNSHTML & "<td width=40% "
		FNSHTML = FNSHTML & "class=ttt  >Instrumento</td>"
		FNSHTML = FNSHTML & "<td width=10% "
		FNSHTML = FNSHTML & "class=ttt  >L.Instrumento</td>"
		FNSHTML = FNSHTML & "<td width=10% "
		FNSHTML = FNSHTML & "class=ttt  >L.Analisis</td></tr>"

		'' FNSHTML = FNSHTML & "<tr><td  "
		'' FNSHTML = FNSHTML & "class=UT colspan=4  >&nbsp;</td></tr>"

		tipocpt="D"
		Strsql="SELECT OBJPRTY.*, Estados.estado as estdes,tipocpt FROM OBJPRTY "
		strsql=Strsql  & " inner join Estados on (OBJPRTY.estado = Estados.objid) "
		strsql=Strsql  & " inner join tipobjeto on (tipobjeto.tipobjeto = OBJPRTY.tipobjeto) "
		Strsql=Strsql & "where OBJPRTY.codobjeto=" & nCodigo

		Set rsx=ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsx.eof Then

			nestado = rsx(9)
			sTipo   = rsx(2)
			tipocpt=rsx(22)
			svienemain = "NO"
			sOrigen1 = "&nbsp;"
			If rsx(3) <> "" Then
				If CLng("0" & rsx(3)) > 0 Then
					svienemain = "SI"
					Strsql = "Select descri From Origen Where objid = " & rsx(3)
					Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If Not rsOrigen.eof Then sOrigen1 = rsOrigen(0)
					Call ofv.cerrarconsulta(rsOrigen)
				End If
			End If


			sInstrumento1 = "&nbsp;"
			If rsx(4) <> "" Then
				Strsql = "Select descri From Instrumen Where objid = " & rsx(4)
				Set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsInstrumento.eof Then sInstrumento1 = rsInstrumento(0)
				Call ofv.cerrarconsulta(rsInstrumento)
			End If

			If rsx(5) <> "" Then
				sInstiden1 =  rsx(5)
			Else
				sInstiden1 = "&nbsp;"
			End If

			If rsx(18) <> "" Then
				wUrl = rsx(18) & ""

				If Not IsNull(wUrl) Then
					slinkD = InStr(wUrl,"://")
					If slinkD > 0 Then
						slinkD = wUrl
					Else
						slinkD = "HTTP://" & wUrl
					End If
				Else
					slinkD = "javascript:void(0);"
				End If

				sInstlink1 = slinkD
				sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & "none" & Chr(34) & ";' href='" & sInstlink1 & "' target=_blank "
				sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
				sInstlink1 = sinput2
			Else
				sInstlink1 = "---"
			End If

			If rsx(19) <> "" Then

				wUrl = rsx(19) & ""

				If Not IsNull(wUrl) Then
					slinkD = InStr(wUrl,"://")
					If slinkD > 0 Then
						slinkD = wUrl
					Else
						slinkD = "HTTP://" & wUrl
					End If
				Else
					slinkD = "javascript:void(0);"
				End If

				sInstlink2 = slinkD


				sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & "none" & Chr(34) & ";' href='" & sInstlink2 & "' target=_blank "
				sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
				sInstlink2 = sinput2
			Else
				sInstlink2 = "---"
			End If

			xTrColor = "#FFFFFF"

			FNSHTML = FNSHTML & "<tr height=22 bgcolor='" & xTrColor & "' >"

			FNSHTML = FNSHTML & "<td  class=ut>"
			'Texto  origen


			SArma = "" & "&nbsp;" & sOrigen1 & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"

			FNSHTML = FNSHTML & "<td  class=ut>"
			'Texto  instrumento

			SArma = "" & sInstrumento1 & "&nbsp;" & sInstiden1 & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"



			FNSHTML = FNSHTML & "<td  class=ut align=center >"
			'link consulta

			SArma = "" & sInstlink1 & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"

			FNSHTML = FNSHTML & "<td  class=ut align=center >"
			'link consulta

			SArma = "" & sInstlink2 & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"

			FNSHTML = FNSHTML & "</tr>"



		End If

		svienemas = "NO"


		Strsql = "Select * From Objprtyadv Where codobjeto=" & nCodigo & " Order By secuen"

		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsx.eof Then   svienemas = "SI"

		Do Until rsx.eof

			If  xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If

			FNSHTML = FNSHTML & "<tr height=22 bgcolor='" & xTrColor & "' >"

			FNSHTML = FNSHTML & "<td  class=ut>"
			'Texto  origen
			sOrigen = ""
			If CStr(rsx(3) & "") <> "" Then
				Strsql = "Select descri From Origen Where objid = " & rsx(3)
				Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsOrigen.eof Then sOrigen = rsOrigen(0)
				Call ofv.cerrarconsulta(rsOrigen)
			End If

			SArma = "" & "&nbsp;" & sOrigen & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"

			FNSHTML = FNSHTML & "<td  class=ut>"
			'Texto  origen
			sInstrumento = ""
			If CStr(rsx(4) & "") <> "" Then
				Strsql = "Select descri From Instrumen Where objid = " & rsx(4)
				Set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsInstrumento.eof Then sInstrumento = rsInstrumento(0)
				Call ofv.cerrarconsulta(rsInstrumento)
			End If
			SArma = "" & sInstrumento & "&nbsp;" & rsx(5) & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"

			If rsx(7) <> "" Then

				wUrl = rsx(7) & ""

				If Not IsNull(wUrl) Then
					slinkD = InStr(wUrl,"://")
					If slinkD > 0 Then
						slinkD = wUrl
					Else
						slinkD = "HTTP://" & wUrl
					End If
				Else
					slinkD = "javascript:void(0);"
				End If

				sInstlink = slinkD
				sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & "none" & Chr(34) & ";' href='" & sInstlink & "' target=_blank "
				sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
				sInstlink = sinput2
			Else
				sInstlink = "---"
			End If

			FNSHTML = FNSHTML & "<td  class=ut align=center >"
			'Texto  identificacion

			SArma = "" & sInstlink & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"

			If rsx(8) <> "" Then

				wUrl = rsx(8) & ""

				If Not IsNull(wUrl) Then
					slinkD = InStr(wUrl,"://")
					If slinkD > 0 Then
						slinkD = wUrl
					Else
						slinkD = "HTTP://" & wUrl
					End If
				Else
					slinkD = "javascript:void(0);"
				End If

				sInstlink = slinkD

				sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
				sinput2 = sinput2  & "none" & Chr(34) & ";' href='" & sInstlink & "' target=_blank "
				sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
				sInstlink = sinput2
			Else
				sInstlink = "---"
			End If

			FNSHTML = FNSHTML & "<td  class=ut align=center >"
			'Texto  identificacion

			SArma = "" & sInstlink & ""
			FNSHTML = FNSHTML & SArma

			FNSHTML = FNSHTML & "</td>"

			FNSHTML = FNSHTML & "</tr>"

			rsx.movenext
		Loop

		Call ofv.cerrarconsulta(rsx)


		'''''''''''''''''''




		FNSHTML = FNSHTML & OFV.genFtabla()
		FNSHTML = FNSHTML & "</center>"
		FNSHTML = FNSHTML & "</DIV>"

		'----------------------------------------------


		'Encabezado empresa

		If SESSION("EMPMUN") = "S" Then

			codempre  =  rs(7)
			codempreL = "N"
			h1        = ""

			If SESSION("EMPMUN") = "S" And SESSION("GLBLOG") = 2 And codempre = 1 And session("un") > 1 Then

				'aca se modifco logo por el del usuario

				codempre = SESSION("CLIEMP")
				codempreL = "&nbsp;&nbsp;<A  href='javascript:void(0);' title='Documento Corporativo.' ><B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>G</FONT></B></a>"
			End If

			'Texto empresa
			Strsql = "Select * From clienteplan where id = " & codempre
			'response.write strsql
			Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsx.eof Then
				sInput1 = rsx(1)
				sInput2 = rsx(5)
			End If

			sinput = "<img  src='../COREss20/Empresa/" & sInput2 & "' >" & "&nbsp;&nbsp;&nbsp;<span style='position:relative;top:-2;' >" & sinput & "</span>"
			h1 = h1 & ofv.GenCelda("","ttt valign=middle ","","1%","","","",sinput,"si")

			sinput = sinput1
			h1 = h1 & ofv.GenCelda("","ttt valign=middle ","","50%","","","",sinput,"si")

			If codempreL <> "N" Then
				h1 = h1 & ofv.GenCelda("","vt ","right","5%","","","",codempreL,"si")
			End If

			fila = ofv.GenRow("","","","","","","",h1,"si")

			FNSHTML = FNSHTML & ofv.GenTabla("","aaa ","","98%","",0,0,2,2,"","",Fila,"si")

		End If


		'FNSHTML = FNSHTML & "<center>"

		FNSHTML = FNSHTML & ofv.GenTabla("","www  style='border-spacing: 8px'","","","",0,0,2,2,"","","","no")


		'Encabezado Descripcion
		h1 = ofv.GenCelda("","tt  ","","20%","","","","ver:","si")
		'Texto descri
		sLink   = "<button value='ver' onclick='chequear();' >"
		h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",  sLink,"si")

		'  FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")





		If Nestado=5 Then


			sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
			sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(0)  & "&ID2=" & rs(0)  & "&ID3=" & rs(0)

			xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
			sHref   = sAccion
			sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
			''    Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","center","","1","","",sInput,"si")
			''  Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")

			Celdaobs= ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
			Celdaobs = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","center","","1","","",Celdaobs,"si")



			'Encabezado Descripcion
			h1 = ofv.GenCelda("","tt  ","","20%","","","","Observaciones:","si")

			h1 = h1 & Celdaobs
			h1 = h1 & ofv.GenCelda("","uti","","80% colspan=2","","","",  "&nbsp;","si")
			FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


		End If







		'Encabezado Descripcion
		h1 = ofv.GenCelda("","tt  ","","20%","","","","Descripcion:","si")
		'Texto descri
		h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",  rs(3),"si")


		'Encabezado Identificacion
		h1 = ofv.GenCelda("","tt  ","","20%","","","","Identificacion:","si")
		'Texto ident
		h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(1),"si")



		'Encabezado Identificacion
		If session("APLMDCLI") = "BPN" Then
			h1 = ofv.GenCelda("","tt  ","","20%","","","","Linea de negocio:","si")
		Else
			'h1 = ofv.GenCelda("","'tt2'  ","","20%","","","","PRODUCTO:","si")
		End If
		'Texto ident
		sInstrumento = "S/D"
		If CStr(rs(9) & "") <> "" Then
			Strsql = "Select descri From TIPOPRODUCTO Where objid = " & rs(9)
			Set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsInstrumento.eof Then sInstrumento = rsInstrumento(0)
			Call ofv.cerrarconsulta(rsInstrumento)
		End If
		if SESSION("APLMDCLI")="BPN" THEN
				h1 = ofv.GenCelda("","'tt2'  ","","60%","","","","<div class=gentit>Linea de Negocio:<div class=gencontent>" & sInstrumento&"</div></div>","si")
				ELSE
				h1 = ofv.GenCelda("","'tt2'  ","","60%","","","","<div class=gentit>Producto:<div class=gencontent>" & sInstrumento&"</div></div>","si")
		END IF
		'h1 = ofv.GenCelda("","'tt2'  ","","60%","","","","<div class=gentit>Producto:<div class=gencontent>" & sInstrumento&"</div></div>","si")
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		h1 = ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>Tema: <div class=gencontent>"& stTema &"</div></div>","si")
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		h1 = ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>Subtema: <div class=gencontent>"& stSubtema &"</div></div>","si")
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		'If sRoot <> "" Then
		h1 = ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>Carpeta: <div class=gencontent>"& sRoot &"</div></div>","si")
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		' end

		''if  sInstrumento <> "S/D" then
		''FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h2,"si")
		''h1=""
		''end if


		'Encabezado Tema
		'h1 = ofv.GenCelda("","tt  ","","20%","","","","Tema:","si")
		'    h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		'Texto Tema
		'h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",stTema,"si")
		'FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

		'Encabezado Subtema
		'h1 = ofv.GenCelda("","tt  ","","20%","","","","SubTema:","si")
		'    h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		'Texto subTema
		'h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",stSubtema,"si")
		'FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

		If sRoot <> "" Then
			'Encabezado ruta
			'h1 = ofv.GenCelda("","tt  ","","20%","","","","Carpeta:","si")
			'   h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
			'Texto ruta
			'h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sRoot,"si")
			'FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

		End If


		sRem = rs(4)
		sRdo = rs(5)

	End If

	Call ofv.cerrarconsulta(rs)



	''''''''''''''''
	'sImagenI3 = "<FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>&#0232;</FONT>"
	sImagenI3 = "<i class='las la-eye' ></i>"
	Strsql="SELECT OBJPRTY.*, Estados.estado as estdes,tipocpt FROM OBJPRTY "
	strsql=Strsql  & " inner join Estados on (OBJPRTY.estado = Estados.objid) "
	strsql=Strsql  & " inner join tipobjeto on (tipobjeto.tipobjeto = OBJPRTY.tipobjeto) "
	Strsql=Strsql & "where OBJPRTY.codobjeto=" & nCodigo

	'Strsql="SELECT OBJPRTY.*, Estados.estado as estdes FROM OBJPRTY, Estados "
	'Strsql=Strsql & "where OBJPRTY.estado = Estados.objid and codobjeto=" & nCodigo

	Set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then

		If svienemas = "SI" Then

			SMASINS = ""


			sinput2 = "<a class=nt  style='color:" & Session("CMFPNBOT") & ";padding-left:0;text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
			sinput2 = sinput2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
			sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "shwOrig('"& rs(0) &"','" & rs(1) & "');" & Chr(34)
			sinput2 = sinput2  & " TITLE='Ver origenes del documento.' >" & sImagenI3 & "&nbsp;VER&nbsp;LISTA&nbsp;...&nbsp;</a>"
			SMASINS = sinput2

		Else

			SMASINS = sOrigen1 & ":&nbsp;" & sInstrumento1 & "&nbsp;" & sInstiden1 & "&nbsp;&nbsp;" & sInstlink1  & "&nbsp;&nbsp;" & sInstlink2  & ""


		End If

		'Encabezado Origen
		'h1 = ofv.GenCelda("","tt  ","","20%","","","","Origen:","si")

		If  svienemain <> "SI" Then
			SMASINS = "S/D"
		End If
		'h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",SMASINS,"si")
		h1 =  ofv.GenCelda("","tt2","","60%","","","","<div class=gentit> Origen: <div class=gencontent>"& SMASINS &"</div></div>","si")
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

		''if  SMASINS <> "S/D" then

		''end if

		'Encabezado Nro de acta
		'h1 = ofv.GenCelda("","tt  ","","20%","","","","Nro.de Acta:","si")
		'combo/texto estado
		'h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(17),"si")

		If  rs(17) & "" <> "" Then
			'		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		End If
		h1 = ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>Nro.de Acta: <div class=gencontent>"& rs(17) &"</div></div>","si")

		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		'Encabezado Estado
		'h1 = ofv.GenCelda("","tt  ","","20%","","","","Estado:","si")
		h1 =  ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>Estado: <div class=gencontent>"& rs(22) &"</div></div>","si")
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

		sVersion = "ORIGINAL"
		If rs(14) = "V" Then
			sVersion = rs(15)
		End If
		h1 = ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>version: <div class=gencontent>"& sVersion &"</div></div>","si")

		'combo/texto estado
		'	h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(22),"si")
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

			Strsql = "Select * From ObjTSEGAGR Where codobjeto=" & nCodigo & " Order By OBJID"
				'response.write strsql
				Set rss = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not RSs.EOF Then

					xFNSHTML = ofv.GenTabla("","","","","",0,0,2,2,"","","","no")
					If session("APLMDCLI") = "BPN" Then
						Celda = ofv.GenCelda("","gentit","","60%","","","","PROCESO:","si")
						Celda = Celda & ofv.GenCelda("","gentit","","60%","","","","SUBPROCESO:","si")
					Else
						Celda = ofv.GenCelda("","","","60%","","","","<div class=gentit>AGRUPAMIENTO</div>","si")

						Celda = Celda & ofv.GenCelda("","","","60%","","","","<div class=gentit>SEGMENTO</div>","si")
					End If
					xFNSHTML = xFNSHTML & ofv.GenRow("","","","","","","",Celda,"si")

					Do Until rss.eof

						sInput = "S/D"
						If CStr(rss(5)&"") <> "" Then
							Strsql = "Select descri From TIPOAGRUPA where OBJID = " & rss(5)
							Set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If Not rsInstrumento.eof Then sInput = rsInstrumento(0)
							Call ofv.cerrarconsulta(rsInstrumento)
						End If

						Celda = ofv.GenCelda("","","","60%","","","","<div class=gencontent>"& sInput & "</div>","si")

						'SEGMENTO

						sInput = "S/D"
						If CStr(rss(4)&"") <> "" Then
							Strsql = "Select descri From TIPOSEGMENTO where OBJID = " & rss(4)
							Set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If Not rsInstrumento.eof Then sInput = rsInstrumento(0)
							Call ofv.cerrarconsulta(rsInstrumento)
						End If

						Celda = Celda & ofv.GenCelda("","","","60%","","","","<div class=gencontent>" & sInput & "</div>","si")

						xFNSHTML = xFNSHTML & ofv.GenRow("","","","","","","",Celda,"si")

						RSs.MOVENEXT
					Loop

					xFNSHTML = xFNSHTML & ofv.GenFTabla()

		End If
             h1 =  ofv.GenCelda("","tt2","","100% ","","","",xFNSHTML,"si")

             FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


		'combo/texto estado
		sRemDes = "S/D"
		If sRem <> "" Then
			Strsql = "Select * From csObjetosRdo Where codobjeto = " & srem
			Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsOrigen.eof Then
				sRemDes = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6) & "(" & rsOrigen(7) & ")"
			End If
			Call ofv.cerrarconsulta(rsOrigen)
		End If
		h1 =  ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>Reemplaza: <div class=gencontent>"& sRemDes &"</div></div>","si")

		'h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sRemDes,"si")

		'' if  sRemDes <> "S/D" then
		FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		'' end if

		strsql1 = "select * from DEPENREEMP where codobj = " & FNCODOBJ
		Set rsOrigen1 = ofv.crearconsultaEx(strsql1,cn,1,parametros)

		Do While Not rsOrigen1.eof
			Strsql = "Select * From csObjetosRdo Where codobjeto = " & rsOrigen1(3)
			Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			If Not rsOrigen.eof Then
				h1 = ofv.GenCelda("","","","20%","","","","","si")
				sInput = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6) & "(" & rsOrigen(7) & ")"
				h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sInput,"si")
				FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
			End If
			Call ofv.cerrarconsulta(rsOrigen)
			rsOrigen1.movenext
		Loop
		Call ofv.cerrarconsulta(rsOrigen1)

		If CInt(rs("estado")) = 4 Then

			'Encabezado Reemplazada
			h1 = ofv.GenCelda("","tt  ","","20%","","","","Reemplazada por:","si")
			'combo/texto estado
			sRemDes = "S/D"
			If srdo <> "" Then
				Strsql = "Select * From csObjetosRdo Where codobjeto = " & srdo
				Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsOrigen.eof Then
					sRemDes = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6) & "(" & rsOrigen(7) & ")"
				End If
				Call ofv.cerrarconsulta(rsOrigen)
			End If
			h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sRemDes,"si")
			h1 =  ofv.GenCelda("","tt2","","60% ","","","","<div class=gentit>Reemplaza por: <div class=gencontent>"& sRemDes &"</div></div>","si")

			''if  sRemDes <> "S/D" then
			FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
			''end if
			'   end if
			'response.write rs("tipocpt")

		End If
		If rs("tipocpt")="K"  Then


			strsql=" select "
			strsql= strsql & " b.descripcion ,"
			strsql= strsql & " c.descripcion ,"
			strsql= strsql & " d.descripcion ,"
			strsql= strsql & "  e.descripcion ,"
			strsql= strsql & "  f.descripcion "
			strsql= strsql & "  from objpropiedades a "
			strsql= strsql & "  left join FORMADMTGUARDA b on ( a.tiempoguarda=b.objid) "
			strsql= strsql & "   left join FORMADMCATG c on ( a.catform=c.objid) "
			strsql= strsql & "   left join FORMADMTFORMU d on ( a.claseform=d.objid) "
			strsql= strsql & "   left join FORMADMCATG e on ( a.catform=e.objid) "
			strsql= strsql & "   left join FORMADMTFIRMA f on ( a.firmasform=f.objid) "
			strsql= strsql & "   where a.codobjeto=" & nCodigo & " and tipobjeto='" & stipo & "'"

			'response.write strsql
			Set rsprop=ofv.crearconsultaEx(StrSql,cn,1,parametros)

			If Not rsprop.eof Then

				h1 = ofv.GenCelda("","tt  ","","20%","","","","TIEMPO GUARDA:","si")
				'combo/texto estado
				sdato = "---"

				If Not isnull(rsprop(0)) Then
					sdato = rsprop(0)
				End If

				h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sdato,"si")
				FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")




				h1 = ofv.GenCelda("","tt  ","","20%","","","","CATEGORIA:","si")

				sdato = "---"

				If Not isnull(rsprop(1)) Then
					sdato = rsprop(1)
				End If

				h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sdato,"si")
				FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


				h1 = ofv.GenCelda("","tt  ","","20%","","","","CLASE&nbsp;FORMULARIO:","si")

				sdato = "---"

				If Not isnull(rsprop(2)) Then
					sdato = rsprop(2)
				End If

				h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sdato,"si")
				FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

				h1 = ofv.GenCelda("","tt  ","","20%","","","","CODIGO&nbsp;FORMULARIO:","si")

				sdato = "---"

				If Not isnull(rsprop(3)) Then
					sdato = rsprop(3)
				End If

				h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sdato,"si")
				FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


				h1 = ofv.GenCelda("","tt  ","","20%","","","","FIRMA&nbsp;FORMULARIO:","si")

				sdato = "---"

				If Not isnull(rsprop(4)) Then
					sdato = rsprop(4)
				End If

				h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sdato,"si")
				FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


			End If
		End If
		' end if

		'Encabezado  Fechas
		h1 = ofv.GenCelda("","ttt colspan=5 ","","20%","","","","Fechas","si")
		'FNSHTML = FNSHTML & ofv.GenRow("","","","","25","","",h1,"si")




		h1 = ""
		h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Creacion","si")
		h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Revision","si")
		h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Publicacion","si")
		h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Vigencia","si")
		h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Baja","si")
		h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Acta","si")
		'FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")








		h1 = ""
		'texto fecha de creacion
		If CLng(rs(6)) = 0 Then
			dFecha = "-"
		Else
			dFecha = CDate(CLng(rs(6)) + CDate("01-01-1900"))
		End If
		h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

		'texto fecha de Vencimiento

		If CLng(rs(13)) = 0 Then
			dFecha = "-"
		Else
			dFecha = CDate(CLng(rs(13)) + CDate("01-01-1900"))
		End If
		h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

		'texto fecha de publicacion

		If CLng( "0" & rs(20)) = 0 Then
			dFecha = "-"
		Else
			dFecha = CDate(CLng(rs(20)) + CDate("01-01-1900"))
		End If
		h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

		'texto fecha de vigencia

		If CLng(rs(7)) = 0 Then
			dFecha = "-"
		Else
			dFecha = CDate(CLng(rs(7)) + CDate("01-01-1900"))
		End If
		h1 = h1 & ofv.GenCelda("","vt style='color:#990000;' ","","","","","",fechadorCor(dFecha,""),"si")

		'texto fecha de eliminacion

		If CLng(rs(8)) = 0 Then
			dFecha = "-"
		Else
			dFecha = CDate(CLng(rs(8)) + CDate("01-01-1900"))
		End If
		h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

		'texto fecha de acta

		If CLng(rs(16)) = 0 Then
			dFecha = "-"
		Else
			dFecha = CDate(CLng(rs(16)) + CDate("01-01-1900"))
		End If
		h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")
		'FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

		stimeline=timeline(rs)


		FNSHTML = FNSHTML & ofv.GenFTabla()





	End If



	GENFMTGENERAL = "<div class='norma-card' id=zglos style='display: grid;grid-template-columns: auto auto;'>" &  FNSHTML  &  stimeline &  "</div>"


End Function


Function timeline(rs)
	timeline="<div class=""right-panel"">"
	timeline=timeline & "<div class=""timeline-title"">Línea de Tiempo</div>"
	timeline=timeline & "      <div class=""timeline"">"
	timeline=timeline & "        <div class=""timeline-event"">"
	timeline=timeline & "          <div class=""event-label"">Creación</div>"
	If CLng(rs(6)) = 0 Then
		dFecha = "-"
	Else
		dFecha = CDate(CLng(rs(6)) + CDate("01-01-1900"))
	End If

	timeline=timeline & "          <div class=""event-date"">" & dFecha & "</div>"
	timeline=timeline & "        </div>"
	timeline=timeline & "        <div class=""timeline-event"">"
	timeline=timeline & "          <div class=""event-label"">Revisión</div>"
	'texto fecha de Vencimiento

	If CLng(rs(13)) = 0 Then
		dFecha = "-"
	Else
		dFecha = CDate(CLng(rs(13)) + CDate("01-01-1900"))
	End If
	timeline=timeline & "         <div class=""event-date"">" & dFecha & " </div>"
	timeline=timeline & "        </div>"
	timeline=timeline & "        <div class=""timeline-event"">"
	timeline=timeline & "          <div class=""event-label"">Publicación</div>"

	If CLng( "0" & rs(20)) = 0 Then
		dFecha = "-"
	Else
		dFecha = CDate(CLng(rs(20)) + CDate("01-01-1900"))
	End If

	timeline=timeline & "          <div class=""event-date"">" & dFecha &"</div>"
	timeline=timeline & "        </div>"
	timeline=timeline & "        <div class=""timeline-event"">"
	timeline=timeline & "          <div class=""event-label"">Vigencia</div>"
	If CLng(rs(7)) = 0 Then
		dFecha = "-"
	Else
		dFecha = CDate(CLng(rs(7)) + CDate("01-01-1900"))
	End If

	timeline=timeline & "          <div class=""event-date alerta"">" & dFecha &"</div>"
	timeline=timeline & "        </div>"
	timeline=timeline & "        <div class=""timeline-event"">"
	timeline=timeline & "          <div class=""event-label"">Baja</div>"

	If CLng(rs(8)) = 0 Then
		dFecha = "-"
	Else
		dFecha = CDate(CLng(rs(8)) + CDate("01-01-1900"))
	End If

	timeline=timeline & "          <div class=""event-date"">" & dFecha &"</div>"
	timeline=timeline & "        </div>"
	timeline=timeline & "        <div class=""timeline-event"">"
	timeline=timeline & "          <div class=""event-label"">Acta</div>"

	If CLng(rs(16)) = 0 Then
		dFecha = "-"
	Else
		dFecha = CDate(CLng(rs(16)) + CDate("01-01-1900"))
	End If
	timeline=timeline & "          <div class=""event-date"">" & dFecha &"</div>"
	timeline=timeline & "        </div>"
	timeline=timeline & "      </div>"
	timeline=timeline & "    </div>"
	timeline=timeline & "  </div>"

End Function

%>