
<%

Function GENFMTGENERALFORM(FNCODOBJ,FNTIPOBJ)


	FNSHTML = ""

	ReDim tobj(0)
	ReDim tobjd(0)

	CLASEFILA="istt"' bgcolor=#f05f23 style='font-weight:bold;color: #ffffff;'"
	CLASEFILAT="ut"' style='font-size:12px;'"
	CLASEFILATF="xxxc style='color: #000000;'"
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

	'response.write stSubtema
	Strsql="Select * From Objetos Where codobjeto=" & nCodigo
	Set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

	If Not rs.eof Then
		sModifica = session("Modificar")
		vuelta= ofv.vueltaasp2("../")

		wdescri ="<i>" & rs(3) & "</i>"

		'----------------------------------------------

		sinput2 = "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=webdings size=1   >r</font>"

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
		FNSHTML = FNSHTML & OFV.gentabla("", "aaa ", "", "98%", "", "0", "", "0", "0", "", "", "", "no")
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
		Strsql="SELECT OBJPRTY.*, Estados.estado as estdes,tipocpt,tipobjeto.descri as tobj,objprty.fecemis,objprty.fecvigen FROM OBJPRTY "
		strsql=Strsql  & " inner join Estados on (OBJPRTY.estado = Estados.objid) "
		strsql=Strsql  & " inner join tipobjeto on (tipobjeto.tipobjeto = OBJPRTY.tipobjeto) "
		Strsql=Strsql & "where OBJPRTY.codobjeto=" & nCodigo
		'response.write strsql
		Set rsx=ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsx.eof Then
			sVersion = "ORIGINAL"
			If rsx(14) &"" = "V" Then
				sVersion = " VERSION (" & rsx(15) & ")"
			End If
			nestado = rsx(9)
			sestado=rsx("estdes")
			sTipo   = rsx(2)
			tipocpt=rsx(22)
			svienemain = "NO"
			sOrigen1 = "&nbsp;"

			If CLng(rsx("fecemis")) = 0 Then
						dFecha = "-"
					Else
						dFecha = CDate(CLng(rsx("fecemis")) + CDate("01-01-1900"))
			End If
		    sfecemis =fechadorCor(dFecha,"")

		    If CLng(rsx("fecvigen")) = 0 Then
			  dFecha = "-"
			Else
			  dFecha = CDate(CLng(rsx("fecvigen")) + CDate("01-01-1900"))
			End If
		    sfecvigen =fechadorCor(dFecha,"")

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

			codempre  = rs(7)
			codempreL = "N"
			h1        = ""

			If SESSION("EMPMUN") = "S" And SESSION("GLBLOG") = 2 And codempre = 1 And session("un") > 1 Then
				codempre = SESSION("CLIEMP")
				codempreL = "&nbsp;&nbsp;<A  href='javascript:void(0);' title='Documento Corporativo.' ><B><FONT FACE=WingDINGS COLOR=#CCCCFF SIZE=5>4</FONT></B></a>"
			End If

			'Texto empresa
			Strsql = "Select * From clienteplan where id = " & codempre
			Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsx.eof Then
				sInput1 = rsx(1)
				sInput2 = rsx(5)
			End If

			sinput = "<img  src=../COREss20/Empresa/" & sInput2 & " >" & "&nbsp;&nbsp;&nbsp;<span style='position:relative;top:-2;' >" & sinput & "</span>"
			h1 = h1 & ofv.GenCelda("","ttt valign=middle ","","1%","","","",sinput,"si")

			sinput = sinput1
			h1 = h1 & ofv.GenCelda("","ttt valign=middle ","","50%","","","",sinput,"si")

			If codempreL <> "N" Then
				h1 = h1 & ofv.GenCelda("","vt ","right","5%","","","",codempreL,"si")
			End If

			fila = ofv.GenRow("","","","","","","",h1,"si")

			FNSHTML = FNSHTML & ofv.GenTabla("","aaa  ","","98%","",0,0,2,2,"","",Fila,"si")

		End If


		FNSHTML = FNSHTML & "<center>"

		FNSHTML = FNSHTML & ofv.GenTabla("","","","98%","",0,0,2,2,"","","","no")


		'Encabezado Descripcion
		h1 = ofv.GenCelda("","tt  ","","20%","","","","ver:","si")
		'Texto descri
		sLink   = "<button value='ver' onclick='chequear();' >"
		h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",  sLink,"si")

		'  FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")





		If Nestado=5 Then


			sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
			sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(0)  & "&ID2=" & rs(0)  & "&ID3=" & rs(0)

			xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
			sHref   = sAccion
			sInput  = "<div style='border-style:solid;border-color:#CCCCFF;border-width:1;' >" & "OBS" & "</div>"
			Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","center","","1","","",sInput,"si")
			Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")


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






		sRem = rs(4)
		sRdo = rs(5)

	End If

	Call ofv.cerrarconsulta(rs)



	''''''''''''''''
	sImagenI3 = "<FONT FACE=WingDINGS COLOR=#f7aa22 SIZE=2>&#0232;</FONT>"


	Strsql="SELECT OBJPRTY.*, Estados.estado as estdes,tipocpt ,tipobjeto.descri as tobj,autor ,isnull(TIPOPRODUCTO.descri,'---') as producto FROM OBJPRTY "
	strsql=Strsql  & " inner join objetos on (OBJPRTY.codobjeto = objetos.codobjeto) "

	strsql=Strsql  & " inner join Estados on (OBJPRTY.estado = Estados.objid) "
	strsql=Strsql  & " inner join tipobjeto on (tipobjeto.tipobjeto = OBJPRTY.tipobjeto) "
	strsql=Strsql  & " left join TIPOPRODUCTO on (TIPOPRODUCTO.objid = objetos.tsegmento) "
	Strsql=Strsql & "where OBJPRTY.codobjeto=" & nCodigo

	'response.write strsql
	Set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then

		If svienemas = "SI" Then

			SMASINS = ""


			sinput2 = "<a class=nt  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
			sinput2 = sinput2  & " underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
			sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "shwOrig('"& rs(0) &"','" & rs(1) & "');" & Chr(34)
			sinput2 = sinput2  & " TITLE='Ver origenes del documento.' >" & sImagenI3 & "&nbsp;VER&nbsp;LISTA&nbsp;...&nbsp;</a>"
			SMASINS = sinput2

		Else

			SMASINS = sOrigen1 & ":&nbsp;" & sInstrumento1 & "&nbsp;" & sInstiden1 & "&nbsp;&nbsp;" & sInstlink1  & "&nbsp;&nbsp;" & sInstlink2  & ""


		End If

		'Encabezado Origen
		h1 = ofv.GenCelda("","tt  ","","20%","","","","Origen:","si")
		If  svienemain <> "SI" Then
			SMASINS = "S/D"
		End If
		h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",SMASINS,"si")


		If  SMASINS <> "S/D" Then
			FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		End If


		'Encabezado Nro de acta
		h1 = ofv.GenCelda("","tt  ","","20%","","","","Nro.de Acta:","si")
		'combo/texto estado
		h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(17),"si")

		If  rs(17) & "" <> "" Then
			FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		End If

		If CLng(rs(7)) = 0 Then
			dFecha = "-"
		Else
			dFecha = CDate(CLng(rs(7)) + CDate("01-01-1900"))
		End If



		estadoDoc="<i>" & rs("estdes") & " </i>"
		tobjdsc="<i>" & rs("tobj") & "</i>"
		sproducto=rs("producto")
		sautor=""

		'Encabezado Reemplazo
		h1 = ofv.GenCelda("","tt  ","","20%","","","","Reemplaza:","si")
		'combo/texto estado
		sRemDes = "S/D"
		If sRem <> "" Then
			Strsql = "Select * From csObjetosRdo Where codobjeto = " & srem
			Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsOrigen.eof Then
				sRemDes = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6)
			End If
			Call ofv.cerrarconsulta(rsOrigen)
		End If
		h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sRemDes,"si")

		If  sRemDes <> "S/D" Then
			FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
		End If

		If CInt(rs("estado")) = 4 Then

			'Encabezado Reemplazada
			h1 = ofv.GenCelda("","tt  ","","20%","","","","Reemplazada por:","si")
			'combo/texto estado
			sRemDes = "S/D"
			If srdo <> "" Then
				Strsql = "Select * From csObjetosRdo Where codobjeto = " & srdo
				Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsOrigen.eof Then
					sRemDes = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6)
				End If
				Call ofv.cerrarconsulta(rsOrigen)
			End If

			h1 =ofv.GenCelda("","uti","","80% colspan=3","","","",sRemDes,"si")



		End If
		If rs("tipocpt")="K" And  (session("CORMODUFORM")="S" Or session("APLMDCLI")="BGAL") Then


			strsql=" select "
			strsql= strsql & " b.descripcion as guarda,"
			strsql= strsql & " c.descripcion  as categoria,"
			strsql= strsql & " d.descripcion as claseform ,"
			strsql= strsql & "  a.codigoform ,"
			strsql= strsql & "  f.descripcion as firmas , "

			strsql= strsql & "  A.MODOIMPRESION ,"
			strsql= strsql & "  A.IMPRESION ,"
			strsql= strsql & "  g.descripcion as COPIAS,"
			strsql= strsql & "  f1.descripcion,A.DESCRIPCION,"
			strsql= strsql & "  g1.descripcion as proceso,"
			strsql= strsql & "  g2.descripcion as BANCA"
			strsql= strsql & "  from objpropiedades a "
			strsql= strsql & "  left join FORMADMTGUARDA b on ( a.tiempoguarda=b.objid) "
			strsql= strsql & "   left join FORMADMCATG c on ( a.catform=c.objid) "
			strsql= strsql & "   left join FORMADMTFORMU d on ( a.claseform=d.objid) "
			'strsql= strsql & "   left join FORMADMCATG e on ( a.codigoform=e.objid) "
			strsql= strsql & "   left join FORMADMTFIRMA f on ( a.firmasform=f.objid) "
			strsql= strsql & "   left join FORMADMPROCESO f1 on ( a.proceso=f1.objid) "
			strsql= strsql & "   left join FORMADMTCOPIAS g on ( a.copias=g.objid) "
			strsql= strsql & "   left join FORMADMPROCESO g1 on ( a.proceso=g1.objid) "
			strsql= strsql & "   left join FORMADMBANCA g2 on ( a.BANCA=g2.objid) "

			strsql= strsql & "   where a.codobjeto=" & nCodigo & " and tipobjeto='" & stipo & "'"
			'response.write strsql
			Set rsprop=ofv.crearconsultaEx(StrSql,cn,1,parametros)

			If Not rsprop.eof Then

				codigoform=rsprop("codigoform")
				h1 = ofv.GenCelda("",CLASEFILA,"","10%","","","","Versión: ","si")
				h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sVersion,"si")


				h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","","Banca: ","si")

				sBANCA="---"
				If Not IsNull(rsprop("BANCA")) Then
					sBANCA=rsprop("BANCA")
				End If

				h1 =  h1 &  ofv.GenCelda("",CLASEFILAT,"","45%","","","",SDATO,"si")
				FNSHTMLP =   ofv.GenRow("","","","","","","",h1,"si")

				h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","","Modo  de Impresion: ","si")
				sMODOIMPRESION = "---"
				If Not IsNull(rsprop("MODOIMPRESION")) Then
					If rsprop("MODOIMPRESION")="SF" Then
						sMODOIMPRESION = "SIMPLE FAZ"
					Else
						sMODOIMPRESION = "DOBLE FAZ"
					End If


				End If
				h1 =  h1 &  ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")

				h1 =  h1 &  ofv.GenCelda("",CLASEFILA,"","","","","","Firmas: ","si")
				sfirmas="---"
				If Not IsNull(rsprop("firmas")) Then

					sfirmas = Replace(rsprop("firmas")," ","&nbsp;")


				End If



			h1 =  h1 &  ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")


			FNSHTMLP =  FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

			sdato = "---"


			h1 =  ofv.GenCelda("",CLASEFILA,"","5%","","","","Tiempo&nbsp;Guarda:","si")

			stguarda = "---"

			If Not IsNull(rsprop(0)) Then
				stguarda = rsprop(0)
			End If

			h1 =h1 & ofv.GenCelda("",CLASEFILAT ,"","40%","","","", sdato ,"si")


			h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","","Producto:","si")

			h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sproducto,"si")
			FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")


			h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Tipo&nbsp;de&nbsp;Formulario:","si")

			sclaseform = "---"

			If Not IsNull(rsprop("claseform")) Then
				sclaseform = rsprop("claseform")
			End If
			h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")


			h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","", "Proceso","si")
			sproceso = "---"

			If Not IsNull(rsprop("proceso")) Then
				sproceso = rsprop("proceso")
			End If

			h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")

			FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

			sdato = "---"
			h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Categoria","si")
			scategoria = "---"

			If Not IsNull(rsprop("categoria")) Then
				scategoria = rsprop("categoria")
			End If

			h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")

			h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","","&nbsp;","si")
			h1 = h1 & ofv.GenCelda("","","","45%","","","","&nbsp;","si")

			FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")
			FNSHTMLP = ofv.GenTabla("","aaa xborder xstyle='border-style:solid;border-width:1;border-color:#f05f23;' cellspacing=2 ","","50%","","0","0","0","0","","",FNSHTMLP,"si")


			'Encabezado Tema

			descri="<b>UBICACION INDICE:<br>" & "Tema: " & stTema & "&nbsp;&nbsp;&nbsp;&nbsp;Subtema: " & stSubtema & "</b>"
			h1 = ofv.GenCelda("",CLASEFILA,"","","","","",descri,"si")


			FNSHTMLC2 = FNSHTMLC2 & ofv.GenRow("","","","","","","",h1,"si")

			TABLATEMA = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23; '","","50%","","0","0","0","0","","",FNSHTMLC2,"si")




			FNSHTMLC2=""


			If Not IsNull(rsprop("DESCRIPCION")) Then
				sdato=Replace(rsprop("DESCRIPCION"),Chr(13),"<BR>")
				'DESCRIPCION
				h1 = ofv.GenCelda("","cc bgcolor=#CCCCFF ","","20%","","","","DESCRIPCION:","si")
				'    h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
				'Texto subTema
				FNSHTMLC2 = FNSHTMLC2 &   ofv.GenRow("","","","","","","",h1,"si")



				h1 = ofv.GenCelda("","vt","","","","","", SDATO  ,"si")
				FILAZ =  ofv.GenRow("","","","","","","",h1,"si")
				TABLADESCRI = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;'","","50%","","0","0","0","0","","",FILAZ,"si")
			End If

			End If





		End If

















		dfecha="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Fecha de Vigencia : <i>" & fechadorCor(dFecha,"") & "</i>"
		' FNSHTML= "<br><TABLE  style='border-style:solid;border-width:1;border-color:black; '  WIDTH=80% cellspacing=2px cellpadding=4px >"
		' FNSHTML=  FNSHTML & " <TR ><TD class=cc xalign=center  COLSPAN=2><font size=2 >"& tobjdsc & "&nbsp;:&nbsp;"  & 	wdescri  &  "</font></TD></TR>"

		'FNSHTML=  FNSHTML & " <TR 	 ><TD class=CC align=center valign=middle COLSPAN=2>CODIGO DEL FORMULARIO :" & sdatoc& "&nbsp;</TD></TR>"
		celda= ofv.GenCelda("",clasefila,"","","","","","Estado :" & estadoDoc & dfecha ,"si")
		FNSHTML = ofv.GenRow("","","","","","","",celda,"si")
		tablaestado  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;' ","","50%","","0","0","0","0","","",FNSHTML,"si")

		'*** Encabezado del Formulario
	'*** BGAL: La descripción ya incluye al código del formulario

	'*** Codigo viejo ***
		'if rs(2) <> "9" then
		'response.write tipocpt
		'If tipocpt = "K" Then
		'  nom= "<h5>" & codigoform & tipocpt & " - " &  wdescri & "</h5>"
		'else
		'   nom= "<h5>" & wdescri & tipocpt & "</h5>"
		'end if

	'*** Codigo nuevo
	'	If session("APLMDCLI")="BGAL" then
		nom= "<h5>" & wdescri  & "</h5>"
	'	Else
	'	nom= "<h5>" & codigoform & " - " &  wdescri & "</h5>"
	'	End If

		' FNSHTML= "<br><TABLE  style='border-style:solid;border-width:1;border-color:black; '  WIDTH=80% cellspacing=2px cellpadding=4px >"
		' FNSHTML=  FNSHTML & " <TR ><TD class=cc xalign=center  COLSPAN=2><font size=2 >"& tobjdsc & "&nbsp;:&nbsp;"  & 	wdescri  &  "</font></TD></TR>"

		'FNSHTML=  FNSHTML & " <TR 	 ><TD class=CC align=center valign=middle COLSPAN=2>CODIGO DEL FORMULARIO :" & sdatoc& "&nbsp;</TD></TR>"
		celda= ofv.GenCelda("",CLASEFILA,"","","","","", nom ,"si")
		FNSHTML = ofv.GenRow("","","","","","","",celda,"si")
		tablanombreform  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;' ","","50%","","0","0","0","0","","",FNSHTML,"si")



		'FNSHTML=    " <TR><TD class=" & clasefila & " xalign=center valign=middle COLSPAN=2>"


		FNSHTML=""

		'  DESCARGAR
		Strsql = "Select DESCRI ,  ARCHIVO ,DIRECTORIO,tipo  From OBJHTM Where codobjeto=" & nCodigo & " Order By OBJID"
		'response.write strsql
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not RS.EOF Then

			FNSHTML =  ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;' ","","50%","",00,0,10,0,"","","","no")
			Celda = ofv.GenCelda("",clasefila & "  COLSPAN=3 ALIGN=CENTER","","","","","","<font size=2> DESCARGA EL FORMULARIO</font>","si")

			FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si")
			POS=1
			Do While Not rs.EOF

				If xTrColor = "#FFFFFF" Then
					xTrColor = "#FFFFFF" '"#F0F0F0"
				Else
					xTrColor = "#FFFFFF"
				End If

				Celda = ofv.GenCelda("","ut","center","5%","","","","&nbsp;","si")

				If CStr(rs(0)&"") <> "" Then
					sInput= "<font size=2> " & Replace(RS(0) &"",Chr(13),"<BR>") & "</font>"
				End If
				Celda = CELDA & ofv.GenCelda("",CLASEFILATF,"","90%","","","", sInput,"si")



				sInput = "S/D"
				If CStr(rs(1)&"") <> "" Then
					sImagenP = "<img width=100px src='imagenes/bajada2.gif'" '"<b><font face=WINGdings size=4 color=#F7AA22>&#052;</font></b>"
					 if session("EMPMUN") = "S" then
      					sarchX="Emp" & SESSION("D" & FNCODOBJ & "-0" )   & "/"
      				 ELSE
      				  sarchX=""
      				 END IF
					If rs("tipo")=3 Then
						sinput = "<a class=nt href='" & RS("archivo")
					Else
						sinput = "<a class=nt href='CoArchivos3.asp?Archivo=" &  sarchX & RS("DIRECTORIO") & "/" & rs(1)
					End If
					sinput = sinput  & "' title='Descargar' target='_blank' >" &  sImagenP & "</a>"

				End If

				Celda = CELDA & ofv.GenCelda("","ut","center","5%","","","",sInput,"si")

				FNSHTML = FNSHTML & ofv.GenRow("","ut  bgcolor='" & xTrColor & "'  ","","","","","",Celda,"si")

				RS.MOVENEXT
				POS=POS +1
			Loop

			FNSHTML =FNSHTML & "</table>"


		End If
	End If


	FNSHTML=  tablanombreform  & "<br><br>" & tablaestado & "<br><br>" & FNSHTML  &  "<br><br>" & FNSHTMLP  & "<br><br>" & TABLATEMA  &"<br><br>" & TABLADESCRI

	FNSHTML = FNSHTML & ofv.GenFTabla()


	FNSHTML=""' <div xclass='container'>"
	FNSHTML= FNSHTML & "    <div class='left-panel'>"
	FNSHTML= FNSHTML & "      <div class='form-title'>"& EsqnomPx(FNTIPOBJ,FNCODOBJ, ofv)& "</div>"
	FNSHTML= FNSHTML & "      <div class='estado-label'>" & sestado &"</div>"


	FNSHTML= FNSHTML & "      <div class='breadcrumb'>"
	FNSHTML= FNSHTML & "        <i class='fas fa-folder-open'></i>"
	FNSHTML= FNSHTML & "        <span>Tema: <strong>" & stTema & "</strong></span>"
	FNSHTML= FNSHTML & "        <i class='fas fa-angle-right'></i>"
	FNSHTML= FNSHTML & "        <span>Subtema: <strong>" & stSubtema & "</strong></span>"
	FNSHTML= FNSHTML & "      </div>"

	FNSHTML= FNSHTML & "      <div class='info-grid'>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Versión</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>" &sVersion & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Banca</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>" & sbanca & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Modo de Impresión</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>" & sMODOIMPRESION & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "       <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Firmas</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>SIN " & sfirma & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Tiempo de Guarda</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>" & stguarda 	& "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Producto</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>"& sproducto & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Tipo de Formulario</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>" & sclaseform & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Proceso</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>" & sproceso & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='info-item'>"
	FNSHTML= FNSHTML & "          <div class='info-label'>Categoría</div>"
	FNSHTML= FNSHTML & "          <div class='info-value'>"&  scategoria & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "      </div>"

	'FNSHTML= FNSHTML & "      <div class='acciones'>"
	'FNSHTML= FNSHTML & "        <button class='button'><i class='fas fa-download'></i> Descargar Formulario</button>"
	'FNSHTML= FNSHTML & "        <button class='button'><i class='fas fa-star'></i> Favorito</button>"
	'FNSHTML= FNSHTML & "        <button class='button'><i class='fas fa-comments'></i> Comentar</button>"
	'FNSHTML= FNSHTML & "      </div>"
	FNSHTML= FNSHTML & "    </div>"

	FNSHTML= FNSHTML & "    <div class='right-panel'>"
	FNSHTML= FNSHTML & "      <div class='timeline-title'>Historial del Formulario</div>"
	FNSHTML= FNSHTML & "      <div class='timeline'>"
	FNSHTML= FNSHTML & "        <div class='timeline-event'>"
	FNSHTML= FNSHTML & "          <div class='event-label'>Creación</div>"
	FNSHTML= FNSHTML & "          <div class='event-date'>" &sfecemis&"</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "        <div class='timeline-event'>"
	FNSHTML= FNSHTML & "          <div class='event-label'>Vigencia</div>"
	FNSHTML= FNSHTML & "          <div class='event-date'>" & sfecvigen & "</div>"
	FNSHTML= FNSHTML & "        </div>"
	FNSHTML= FNSHTML & "      </div>"
	FNSHTML= FNSHTML & "    </div>"
   '  FNSHTML= FNSHTML & "</div>"

	GENFMTGENERALFORM = "<div id=zglos class='norma-card' style='display: grid;grid-template-columns: auto auto;' >" &  FNSHTML  & "</div>"


End Function




%>