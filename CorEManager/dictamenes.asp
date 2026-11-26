<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	shtml= shtml & "<body style='background-color:#E0F2F7;'>"
	response.expires=0

	CantRegMostrarx = 11
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	CODIGO=REQUEST.QUERYSTRING("CODIGO")
	TIPO=REQUEST.QUERYSTRING("TIPO")
	ACC=REQUEST.QUERYSTRING("ACC")
	pend=REQUEST.QUERYSTRING("pend")
	objid=request.querystring("objid")
	Set cn = ofv.conectar(ofv.strconn5)
	'response.write CODIGO
	'response.write request.querystring

	strsql = "select segckt from  segcircuito  where OBJCOD='" & CODIGO & "' and objtipo='" & TIPO & "' and model in ('WKF00016631121722600921928', 'WKF00016631121722600921928DD')"
	'model='WKF00016631121722600921928' order by id desc"
	'response.write strsql & "<br>"
	Set tt = ofv.crearconsultaex(strsql,cn,1,parametros)
	segcktfin =""

	If Not tt.eof Then
		segcktfin = tt(0)
	End If
	Call ofv.cerrarconsulta(tt)

	If segcktfin <> "" Then
		sql = "select a.respasg,a.observaciones,a.pscfin,a.responsable,a.pscdsc,b.segsalida from cktview a "
		sql = sql & " inner join segpaso b on (a.segckt = b.segckt and a.segcktpsc = b.segcktpsc) "
		sql = sql & " where a.segckt='" & segcktfin & "' and a.pscdsc <> '10: INICIO' order by pscfin "
		'response.write sql
		Set rs = ofv.crearconsultaex(sql,cn,1,parametros)


		celda = ofv.GenCelda("", "ttt align=center style='background-color:#77AADD;font-family:Tahoma;font-size:14' ;", "", "", "", "", "", "<b>OBSERVACIONES DEL WORKFLOW</b>", "si")
		fila  = ofv.GenRow("", "", "", "", "", "", "",celda, "si")
		sHTML = sHTML & ofv.GenTabla("", "AAA STYLE='BORDER-STYLE:ridge;BORDER-WIDTH:2;border-color:#f7aa22;' ", "", "100%", "", "0", "", "0", "0", "", "", fila, "si")  &"<br>"

		Do While Not rs.eof
			celda=""
			fila=""
			rechazado=""

			If (rs(4) <> "999:FINAL" And rs(4) <> "30:RECHAZO DE RESPONSABLES") Then

				'if instr(rs(1) & "" ,"<RCH>") > 0 then rechazado = "(Rechazado)"
				If rs("segsalida") = "RCH" Then rechazado = "(Rechazado)"

				If rs(0) = rs(3) Then
					xxxx1     = getInfoUser(rs(0))
					xusrlog1  = session("SGU_nombre")
					xusrlog1 = xusrlog1 & "&nbsp;&nbsp;&nbsp;&nbsp;(ORIGINANTE)"
				Else
					xxxx1     = getInfoUser(rs(0))
					xusrlog1  = session("SGU_nombre")
				End If

				celda = celda & ofv.GenCelda("", "ttt style='font-family:Tahoma;background-color:#87CEFA;font-size:12';", "", "", "", "", "",xusrlog1 & "&nbsp;&nbsp;&nbsp;&nbsp;" & rechazado, "si")
				celda = celda & ofv.GenCelda("", "ttt align=right style='font-family:Tahoma;background-color:#87CEFA;font-size:12';", "", "", "", "","",rs(2), "si")
				fila  = ofv.GenRow("", "", "", "", "10", "", "",celda, "si")

				If rs(1)<> "" Then
					celda = ofv.GenCelda("", "ttt colspan=2 style='font-family:Tahoma;background-color:#F7EEDD;font-size:12';", "", "", "", "", "",rs(1),"si")
				Else
					celda = ofv.GenCelda("", "ttt colspan=2 style='font-family:Tahoma;background-color:#F7EEDD;font-size:12';", "", "", "", "", "","<font color=#FF0000><i>Sin observaciones al momento</i></font>", "si")
				End If
				fila  = fila & ofv.GenRow("", "", "", "", "10", "", "",celda, "si")

				sHTML = sHTML & ofv.GenTabla("", "AAA STYLE='BORDER-STYLE:ridge;BORDER-WIDTH:2;border-color:#f7aa22;' ", "", "100%", "", "0", "", "0", "0", "", "", fila,"si") & "<br>"
			End If
			rs.movenext
		Loop
	End If

	Call ofv.cerrarconsulta(rs)


	sHTML = sHTML & "</body></html>"
End If


Response.Write sHTML
Call ofv.cerrarconn(cn)

%>