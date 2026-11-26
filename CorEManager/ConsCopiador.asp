<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires = 0
	server.scripttimeout = 600
	xx=getinfouser(session("usuario"))
	codTEMA=CStr(request.querystring("codTEMA")&"")
	'tipoTEMA=request.querystring("tipoTEMA")

	codSUBTEMA=request.querystring("codSUBTEMA")
	tipoSUBTEMA=request.querystring("tipoSUBTEMA")

	codobjeto=request.querystring("codobjeto")
	tipobjeto=request.querystring("tipobjeto")
	tid=request.querystring("tid")
	idDoc=request.querystring("id")

  ' ON ERROR RESUME NEXT
	'response.write request.querystring & "<br>"
	WSPDIVN = "DTTEMAINT"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
	sHTML = ""

	sHTML = sHTML & "<HTML>"

	sHTML = sHTML & "<BODY onload=""window.top.document.getElementById('DTTEMAINT').innerHTML=window.top.document.getElementById('xxxx').contentWindow.document.body.innerHTML;window.top.document.getElementById('DTTEMAINT').style.visibility='visible';"" >"
	''''''''''''''


	XXXX = NEWEMPRE("NO")


	Set cn = OFV.conectar(OFV.strconn2)

	sql="select cliente,archivo  from clienteplan where id=" & SESSION("UN")
	Set rsTipo = OFV.crearconsultaEx(sql, cn, 1, volver)

	If Not rstipo.eof Then
		scliente= rstipo(0)
		sclogo= rstipo(1)
	End If

	Call OFV.cerrarconsulta(rsTipo)

	Select Case  WKAPLIC
		Case "EM"
		WHERETIPO= " WHERE APLIC IN ('EM','GLB')"
		If session("base")="K" Then
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT = 'K'"
		Else
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT <> 'K'"
		End If
		Case "GLB"
		WHERETIPO= " WHERE APLIC IN ('EM','GLB')"
		'WHERETIPO2= " AND APLIC IN ('EM','GLB')"
		If session("base")="K" Then
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT = 'K'"
		Else
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT <> 'K'"
		End If
		Case Else
		WHERETIPO=""
		WHERETIPO2=""
	End Select
	APLCHECK()


	strsql = "Select nivel From Tipobjeto Where tipobjeto = '" & tid & "'"

	If session("base")="K" Then
		ADDW=" ICONO='K' AND "
	Else
		ADDW=" ICONO <> 'K' AND "
	End If

	Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
	sTipo = rsTipo(0)
	Call OFV.cerrarconsulta(rsTipo)
	Select Case sTipo
		Case "C"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levc = 1"
		Case "D"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levd = 1"
		Case "E"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "leve = 1"
		Case "F"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levf = 1"
		Case "G"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levg = 1"
		Case "H"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levh = 1"
		Case "I"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levi = 1"
		Case "J"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levj = 1"
		Case "K"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levk = 1"
		Case "L"
		strsql = "Select tipobjeto From Subtema Where " & ADDW & "levl = 1"
	End Select
	' response.write strsql
	Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
	aTipos0 = "('"
	Do Until rsTipo0.EOF
		aTipos0 = aTipos0 & rsTipo0(0) & "','"
		rsTipo0.movenext
	Loop
	aTipos0 = Left(aTipos0, Len(aTipos0) - 2) & ")"

	Select Case sTipo
		Case "C"
		strsql = "Select tipobjeto From Tipobjeto Where levc = 1 " & wheretipo2
		Case "D"
		strsql = "Select tipobjeto From Tipobjeto Where levd = 1" & wheretipo2
		Case "E"
		strsql = "Select tipobjeto From Tipobjeto Where leve = 1" & wheretipo2
		Case "F"
		strsql = "Select tipobjeto From Tipobjeto Where levf = 1" & wheretipo2
		Case "G"
		strsql = "Select tipobjeto From Tipobjeto Where levg = 1" & wheretipo2
		Case "H"
		strsql = "Select tipobjeto From Tipobjeto Where levh = 1" & wheretipo2
		Case "I"
		strsql = "Select tipobjeto From Tipobjeto Where levi = 1" & wheretipo2
		Case "J"
		strsql = "Select tipobjeto From Tipobjeto Where levj = 1" & wheretipo2
		Case "K"
		strsql = "Select tipobjeto From Tipobjeto Where levk = 1" & wheretipo2
		Case "L"
		strsql = "Select tipobjeto From Tipobjeto Where levl = 1" & wheretipo2
	End Select
	Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
	aTipos = "('"
	Do Until rsTipo.EOF
		aTipos = aTipos & rsTipo(0) & "','"
		rsTipo.movenext
	Loop

	aTipos = Left(aTipos, Len(aTipos) - 2) & ")"
	'response.write aTipos & "<br>" & aTipos0


	SQL ="select distinct codtema,DESCRI from tema "
	'sql3=SQL
	if Session("EMPMUN") = "S" then
	 SQL=SQL & " inner join ATRMTEMAST on (ATRMTEMAST.atrmcodst=tema.codtema and atrmcodes=" & SESSION("UN") & " and atrmverst=1 and atrmtipost='T') "
	end if
	 Strsql = strsql & " inner join prodserv b on (tema.codtema = b.codtema and b.marcaps <> 'PEDIDO' and b.marcaps <> 'SNN') "
	 sql= sql & " WHERE "
     sand=""
	'RESPONSE.WRITE codTEMA & "222"
	If codTEMA <> "0" And codTEMA <> "" Then
		SQL=SQL & "  tema.codtema=" & codTEMA
		sand=" and "
		'SQL3=sql3 &  " where tema.codtema=" & codTEMA
	End If

	If session("base")="K" Then
		SQL= SQL & sand & " tema.PSTIPO='K' "
		sand=" and "
	Else
		SQL=SQL  & sand & "  tema.PSTIPO <> 'K' "
		sand=" and "
		'SQL3=sql3 &  " 1=1 AND tema.PSTIPO <> 'K' "
	End If


	SQL=SQL & " ORDER BY DESCRI"
	'SQL3=sql3  & " ORDER BY DESCRI"
   ' response.write sql
	shtml= shtml & "<table  width=100% >"
	shtml= shtml & "<tr>"
	shtml= shtml & "<td colspan=10>"
	IF Session("EMPMUN") = "S" THEN
	shtml= shtml & "<a class=ut href=conscopiador.asp?id=" & idDoc & "&tid=" & tid &"&codtema=0 target=xxxx><b>TEMAS :" &scliente & "</b></a>"
	ELSE
	shtml= shtml & "<a class=ut href=conscopiador.asp?id=" & idDoc & "&tid=" & tid &"&codtema=0 target=xxxx><b>TEMAS</b></a>"
	END IF

	shtml= shtml & "</td>"
	shtml= shtml & "</tr>"





	'response.write sql

	Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

	Do While Not rs.eof
		shtml= shtml & "<tr>"
		shtml= shtml & "<td width=1% >&nbsp;</td><td colspan=9>"
		shtml=shtml & "<a class=ut href=conscopiador.asp?id=" & idDoc & "&tid=" & tid & "&codtema=" & rs(0) & " target=xxxx>TEMA: " & rs(1) & "</a>"
		shtml= shtml & "</td>"
		shtml= shtml & "</tr>"
		'  EXIT DO
		rs.movenext
	Loop





	If codTEMA <> "0" And codTEMA <> "" Then

		sql =" select codprodserv,marcaps,descri from prodserv "
		sql = sql & "  where codtema=" & codtema &" and marcapS in " & atipos0

		If codsubtema <> "0" And codsubtema <> "" Then
			sql = sql & " AND codprodserv=" &  codsubtema
		End If

		If session("base")="K" Then
			SQL= SQL & " AND prodserv.PSTIPO='K' "
		Else
			SQL=SQL & " AND prodserv.PSTIPO <> 'K' "
		End If

		SQL= SQL &" ORDER BY DESCRI"
		'response.write sql
		Call ofv.cerrarconsulta(rs)
		Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

		Do While Not rs.eof
			shtml= shtml & "<tr>"
			shtml= shtml & "<td width=1% >&nbsp;</td><td width=1% >&nbsp;</td><td colspan=8>"
			Sasignador="<span style='display:inline-block; width: 10;'></span> <font title='ASIGNAR' style='cursor: pointer; cursor: hand;' font face=corflaticon size=5 onclick=asignar('" & rs(0) & "_" & rs(1) & "') >a</font>"
			shtml=shtml  & "<a class=ut id='" & rs(0) & "_" & rs(1) & "'  href=conscopiador.asp?id=" & idDoc & "&tid=" & tid & "&codtema=" & codTEMA & "&codsubtema=" & rs(0) & "&tiposubtema=" & rs(1) &" target=xxxx>PRODUCTO: "   & rs(2) & "</a>"
			shtml =shtml & Sasignador
			shtml= shtml & "</td>"
			shtml= shtml & "</tr>"
			rs.movenext
		Loop


		sql = "SELECT DISTINCT OBJETOS.codobjeto,OBJETOS.tipobjeto,OBJETOS.descri,tipobjeto.tipocpt,estados.estado FROM RELAOBJ, OBJETOS, OBJPRTY,"
		sql = sql & " Estados,tipobjeto where RELAOBJ.codojasoc = OBJETOS.codobjeto and OBJETOS.tipobjeto=tipobjeto.tipobjeto and OBJPRTY.estado =estados.objid and OBJPRTY.estado <> 4 and OBJETOS.codobjeto = OBJPRTY.codobjeto and (objetos.objun = " & session("un") & " or objetos.objun = 1 ) "
		sql = sql & "  And Objetos.tipobjeto In " & atipos
		sql = sql & " and OBJETOS.codobjeto <> " & idDoc
		'response.write sql
		If codobjeto  <> "0" And codobjeto <> "" Then
			sql=sql & " and RELAOBJ.codobjmaster =0" & codobjeto & " and relaobj.tipobjmaster = '"& tipobjeto  &"'"

			sqlpath="select mc1.descri,mc2.descri,mc3.descri,mc4.descri,mc5.descri,mc6.descri,mc7.descri,a.* from  csrelpath a "
			sqlpath= sqlpath & " left join objetos mc1 on(mc1.codobjeto=a.mc1 and mc1.tipobjeto=a.mt1 )"
			sqlpath= sqlpath & "left join objetos mc2 on(mc2.codobjeto=a.mc2 and mc2.tipobjeto=a.mt2 ) "
			sqlpath= sqlpath & "left join objetos mc3 on(mc3.codobjeto=a.mc3 and mc3.tipobjeto=a.mt3 ) "
			sqlpath= sqlpath & "left join objetos mc4 on(mc4.codobjeto=a.mc4 and mc4.tipobjeto=a.mt4 ) "
			sqlpath= sqlpath & "left join objetos mc5 on(mc5.codobjeto=a.mc5 and mc5.tipobjeto=a.mt5) "
			sqlpath= sqlpath & "left join objetos mc6 on(mc6.codobjeto=a.mc6 and mc6.tipobjeto=a.mt6) "
			sqlpath= sqlpath & "left join objetos mc7 on(mc7.codobjeto=a.mc7 and mc7.tipobjeto=a.mt7 )"
			sqlpath= sqlpath & " where a.cobj=0"

		Else
			sqlpath=""
			sql=sql & " and RELAOBJ.codobjmaster =0" & codsubtema & " and relaobj.tipobjmaster = '"& tiposubtema  &"'"
		End If

		Set rso = ofv.crearconsultaEx(sql,cn,1,parametros)
		celda="<td width=1% >&nbsp;</td><td width=1% >&nbsp;</td><td width=1% >&nbsp;</td>"
		If Not rso.eof And sqlpath <>"" Then

			sqlpath= sqlpath & rso(0)

			' RESPONSE.WRITE sqlpath
			Set rspth=ofv.crearconsultaEx(sqlpath,cn,1,parametros)
			If Not rspth.eof Then

				POS=7
				posF=1
				For i=0 To 6
					If Trim(rspth(i) & "") <> "" Then
						shtml= shtml & "<tr>"
						SHTML=SHTML & CELDA
						SHTML=SHTML & "<TD COLSPAN=" & POS & " >"
						xtipo="mt" & (posf + i)
						xcod= "mc" & (posf + i)
						Sasignador="<span style='display:inline-block; width: 10;'></span> <font title='ASIGNAR' style='cursor: pointer; cursor: hand;' font face=corflaticon size=5 onclick=asignar('" &  rspth(xcod) & "_" & rspth(xtipo) & "') >a</font>"
						shtml=shtml  & "<a class=ut  id="  & rspth(xcod) & "_" & rspth(xtipo) & "  href=conscopiador.asp?id=" & idDoc & "&tid=" & tid & "&codobjeto=" &"" & "&tipobjeto="& "" & "&codtema=" & codTEMA & "&codsubtema=" & codsubtema & "&tiposubtema=" & tiposubtema &" target=xxxx>"   & EsqnomPx(rspth(xtipo),rspth(xcod),ofv)  & "</a>"
						shtml= shtml & Sasignador &"</td>"
						shtml= shtml & "</tr>"
						CELDA=CELDA & "<td width=1% >&nbsp;</td>"
						shtml= shtml & "</tr>"
						POS=POS-1
					End If
				Next
			End If
		End If


		Do While Not rso.eof
			shtml= shtml & "<tr>"
			shtml=shtml  & CELDA
			SHTML=SHTML & "<TD COLSPAN=" & POS & " >"
			If rso("tipocpt") = "C" Then
				sdato=EsqnomPx(rso("tipobjeto"),rso("codobjeto"),ofv)
				Sasignador="<span style='display:inline-block; width: 10;'></span> <font title='ASIGNAR' style='cursor: pointer; cursor: hand;' font face=corflaticon size=5 onclick=asignar('" &  rso(0) & "_" & rso(1) & "') >a</font>"
				shtml=shtml  & "<a  class=ut id='" & rso(0) & "_" & rso(1) & "' xonclick=asignar('" & rso(0) & "_" & rso(1) & "')  href=conscopiador.asp?id=" & idDoc & "&tid=" & tid & "&codobjeto=" & rso(0) & "&tipobjeto="& rso(1) & "&codtema=" & codTEMA & "&codsubtema=" & codsubtema & "&tiposubtema=" & tiposubtema &" target=xxxx>"   & EsqnomPx(rso("tipobjeto"),rso("codobjeto"),ofv) & "</a>"
			Else
				Sasignador="<span style='display:inline-block; width: 10;'></span> <font title='ASIGNAR' style='cursor: pointer; cursor: hand;' font face=corflaticon size=5 onclick=asignar('" &  rso(0) & "_" & rso(1) & "') >a</font>"
				shtml=shtml  &  "<a class=ut  id='" & rso(0) & "_" & rso(1) & "' shref=javascript:asignar('"& rso(0) & "_" & rso(1)& "'); >" & EsqnomPx(rso("tipobjeto"),rso("codobjeto"),ofv) & "    (" & rso("estado") & ")" & "</a>"
			End If
			shtml= shtml & Sasignador & "</td>"
			shtml= shtml & "</tr>"
			rso.movenext
		Loop



	End If


	SHTML=SHTML & "</table>"

	jsc =   Chr(13) & "<SCRIPT language=javascript >" & Chr(13)
	jsc = jsc & Chr(13)
	jsc = jsc & "function asignar(e) {" & Chr(13)
	jsc = jsc & " alert('1');"
	jsc = jsc & " } " & Chr(13)
	jsc = jsc & " </SCRIPT> " & Chr(13)

	sHTML = sHTML '& jsc
	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


	sHTML = sHTML & "</body></html>"




	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	'''''''''''''

End If

response.write sHTML





%>
