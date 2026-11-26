<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
APLCHECK()
Sub linea(xfila,xindi)
	If xindi > 50 Then
		xindi = 0
		response.write "</table>"
		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"

		If ofv.explorador = "MSIE" Then response.write "<div align=center > &nbsp </div>"

		response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>" & Chr(13)
		response.write "<tr>"
		response.write "<td width=100% >"
		response.write ofv.RepHeader2("Lista de Esquemas Normativos Vigentes","600")

		response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
		response.write "<tr>"
		response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
		response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
		response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
		response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
		response.write "<td width=5>&nbsp;</td>"
		response.write "</tr>"
	End If

	response.write xfila
	xindi = xindi + 1
End Sub

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires = 0
	Server.ScriptTimeout=60 * 10
	desde = request.querystring("desde")
	hasta = request.querystring("hasta")
	tipo = request.querystring("tipo")

	''  response.write ofv.menuHeader("","")

	response.write CorStyler()
	If SESSION("BASE")="K" Then
		response.write  ofv.RepHeader2("Lista de Formularios Vigentes","600")
	Else

		response.write  ofv.RepHeader2("Lista de Esquemas Normativos Vigentes","600")
	End If

	styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"






	indi = 0
	response.write "<table  width=100% marginwidth=0  cellspacing=0 cellpadding=0>"
	response.write "<tr>"
	response.write "<td  align=center >"
	response.write "<table id=margen1 width=95% marginwidth=0  cellspacing=0 cellpadding=2  >"
	response.write "<tr>"
	response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
	response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
	response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
	response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
	response.write "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
	response.write "</tr>"

	Set cn = ofv.conectar(ofv.strconn2)

	'xx Arma tipos de objeto
	Dim tobj()
	Dim tobjd()
	inx = 0
	ReDim Preserve tobj(0)
	ReDim Preserve tobjd(0)

	Strsql = "Select * from tipobjeto "

	If tipo <> "TODOS" Then

		Strsql = strsql & " where TIPOCPT='C'  OR tipobjeto='" & tipo & "' "

	End If

	Strsql = strsql & " order by tipobjeto"
	'response.write strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	SQLTIPo="("
	DELIM=""
	Do Until rs.eof
		SQLTIPo =SQLTIPo & DELIM & "'" & RS(0) & "'"
		DELIM=","
		ReDim Preserve tobj(inx)
		ReDim Preserve tobjd(inx)
		tobj(inx) = rs(0)
		tobjd(inx) = rs(1)
		inx = inx + 1
		rs.movenext
	Loop
	SQLTIPO=SQLTIPO &")"

	'xx Arma tipos de subtema
	Dim tsub()
	Dim tsubd()
	inx = 0
	ReDim Preserve tsub(0)
	ReDim Preserve tsubd(0)

	If session("EMPMUN") = "S" Then
		Strsql = "Select distinct a.* from subtema a "
		strsql = strsql & " inner join atrmtemast b on (b.atrmtipost='S' and b.atrmcodes=" & session("UN") & " and b.atrmverst=1) "
		strsql = strsql & " order by a.tipobjeto"
	Else
		Strsql = "Select * from subtema order by tipobjeto "
	End If
	'response.write strsql & "<br>"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	Do Until rs.eof
		ReDim Preserve tsub(inx)
		ReDim Preserve tsubd(inx)
		tsub(inx) = rs(0)
		tsubd(inx) = rs(1)
		inx = inx + 1
		rs.movenext
	Loop
	Call ofv.cerrarconsulta(rs)
	' Fechas tope

	nFecini = 0

	If nFecini < CLng(CDate("01-01-1900")) Then
		nFecini = CLng(CDate("01-01-1900"))
	Else
		nFecini = 0
	End If

	xfnul = CLng(CDate("01-01-1900")) - nFecini
	xfhoy = CLng(Date()) - nFecini

	'xx Arma Temas:
	If session("EMPMUN") = "S" Then
		Strsql = "Select distinct a.* from Tema a "
		strsql = strsql & " inner join atrmtemast b on (b.atrmtipost='T' and b.atrmcodes=" & session("UN") & " and b.atrmverst=1) "
		If SESSION("BASE")="K" Then
			Strsql = " WHERE a.PSTIPO='K' "
		Else
			Strsql = " WHERE a.PSTIPO <> 'K' "
		End If

		strsql = strsql & " order by a.descri"


	Else
		'Strsql = "Select * from Tema order by descri"
		If SESSION("BASE")="K" Then
			Strsql = "Select * from Tema WHERE PSTIPO='K' order by descri"
		Else
			Strsql = "Select * from Tema WHERE PSTIPO <> 'K' order by descri"
		End If
	End If
	'response.write strsql & "<br>"
	Set rsa = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	Do Until rsa.eof
		temax = "<tr height=30 ><td colspan=14  class=ttt >TEMA: " & rsa(1) & "</td></tr>"
		'xx Arma SubTemas:
		Strsql = "Select * from prodserv where codtema = " & rsa(0) & " order by descri"
		Set rsb = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		Do Until rsb.eof
			tipox = "SUBTEMA: "
			For inx=0 To UBound(tsub)
				If tsub(inx) = rsb(4) Then
					tipox = tsubd(inx) & ": "
					Exit For
				End If
			Next
			subtemax = "<tr><td width=5>&nbsp;</td>"
			subtemax = subtemax & "<td colspan=13  class=vt >" & tipox & rsb(1) & "</td></tr>"
			'xx Arma Objetos C:
			Set cnc = ofv.conectar(ofv.strconn2)

			Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos a"
			Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
			Strsql = Strsql & " inner join objetos c on (a.codobjeto = c.codobjeto and a.tipobjeto=c.tipobjeto) "
			Strsql = Strsql & " where a.codobjmaster=" & rsb(0)
			Strsql = Strsql & " and a.tipobjmaster='" & rsb(4)
			Strsql = Strsql & "' And a.estado = 3 "

			If tipo <> "TODOS" Then
				Strsql = Strsql & " and a.tipobjeto IN" & SQLTIPO
			End If

		'	If session("EMPMUN") = "S" Then
				Strsql = Strsql & " And c.objun =" & session("UN")
			'End If



			if desde = "TODOS"  then
						else
						Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
						end if


						if  hasta = "TODOS" then
						Strsql = Strsql & ""
						else
						 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
						end if


			Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
			'response.write strsql & "<br>"
			Set rsc = ofv.crearconsultaEx(StrSql,cnc,1,parametros)
			If Not rsc.eof Then
				Call linea(temax,indi)
				Call linea(subtemax,indi)
				temax = "<tr><td colspan=13  class=vt >&nbsp;</td></tr>"
			End If
			Do Until rsc.eof
				tipox = "Esquema: "
				For inx=0 To UBound(tobj)
					If tobj(inx) = rsc(2) Then
						tipox = tobjd(inx) & ": "
						Exit For
					End If
				Next

				subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
				'subtemax = subtemax & "<td colspan=11  class=ut >" & tipox & rsc(1) & "</td>"
				If session("APLMDCLI")="BC" Then
					subtemax = subtemax & "<td colspan=11  class=ut >" & tipox & rsC("CODIDEN") & " : " & RSC(1) & "</td>"
				Else
					subtemax = subtemax & "<td colspan=11  class=ut >" & tipox & rsC(1) & "</td>"
				End If

				dvige = CStr(rsc(4) & "")
				If CLng("0" & dvige) > 0 Then
					subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
				Else
					subtemax = subtemax & "<td class=ut >&nbsp;</td>"
				End If



				subtemax = subtemax & "</tr>"

				Call linea(subtemax,indi)
				'xx Arma Objetos D:
				Set cnd = ofv.conectar(ofv.strconn2)
				Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
				Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
				Strsql = Strsql & " where a.codobjmaster=" & rsc(0)
				Strsql = Strsql & " and a.tipobjmaster='" & rsc(2)
				Strsql = Strsql & "' And a.estado = 3 "

				If tipo <> "TODOS" Then
					Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
				End If

			'	If session("EMPMUN") = "S" Then
					Strsql = Strsql & " And a.objun =" & session("UN")
			'	End If

			if desde = "TODOS"  then
						else
						Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
						end if


						if  hasta = "TODOS" then
						Strsql = Strsql & ""
						else
						 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
						end if

				Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"

				Set rsd = ofv.crearconsultaEx(StrSql,cnd,1,parametros)
				Do Until rsd.eof
					tipox = "Esquema: "
					For inx=0 To UBound(tobj)
						If tobj(inx) = rsd(2) Then
							tipox = tobjd(inx) & ": "
							Exit For
						End If
					Next
					subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
					'subtemax = subtemax & "<td colspan=10  class=ut >" & tipox & rsd(1) & "</td>"
					If session("APLMDCLI")="BC" Then
						subtemax = subtemax & "<td colspan=10  class=ut >" & tipox & rsD("CODIDEN") & " : " & RSD(1) & "</td>"
					Else
						subtemax = subtemax & "<td colspan=10  class=ut >" & tipox & rsd(1) & "</td>"
					End If

					''    subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsd(4) + nFecini) & "</td>"

					dvige = CStr(rsd(4) & "")
					If CLng("0" & dvige) > 0 Then
						subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
					Else
						subtemax = subtemax & "<td class=ut >&nbsp;</td>"
					End If

					subtemax = subtemax & "</tr>"
					Call linea(subtemax,indi)
					'xx Arma Objetos E:
					Set cne = ofv.conectar(ofv.strconn2)
					Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
					Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
					Strsql = Strsql & " where a.codobjmaster=" & rsd(0)
					Strsql = Strsql & " and a.tipobjmaster='" & rsd(2)
					Strsql = Strsql & "' And a.estado = 3 "

					If tipo <> "TODOS" Then
						Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
					End If

				'	If session("EMPMUN") = "S" Then
						Strsql = Strsql & " And a.objun =" & session("UN")
				'	End If


				if desde = "TODOS"  then
							else
							Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
							end if


							if  hasta = "TODOS" then
							Strsql = Strsql & ""
							else
							 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
							end if


					Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
					Strsql = Strsql & "  "
					Set rse = ofv.crearconsultaEx(StrSql,cne,1,parametros)
					Do Until rse.eof
						tipox = "Esquema: "
						For inx=0 To UBound(tobj)
							If tobj(inx) = rse(2) Then
								tipox = tobjd(inx) & ": "
								Exit For
							End If
						Next
						subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
						subtemax = subtemax & "<td width=5>&nbsp;</td>"
						'subtemax = subtemax & "<td colspan=9  class=ut >" & tipox & rse(1) & "</td>"
						If session("APLMDCLI")="BC" Then
							subtemax = subtemax & "<td colspan=9 class=ut >" & tipox & rsE("CODIDEN") & " : " & RSE(1) & "</td>"
						Else
							subtemax = subtemax & "<td colspan=9  class=ut >" & tipox & rsE(1) & "</td>"
						End If

						''   subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rse(4) + nFecini) & "</td>"

						dvige = CStr(rse(4) & "")
						If CLng("0" & dvige) > 0 Then
							subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
						Else
							subtemax = subtemax & "<td class=ut >&nbsp;</td>"
						End If


						subtemax = subtemax & "</tr>"
						Call linea(subtemax,indi)
						'xx Arma Objetos F:
						Set cnf = ofv.conectar(ofv.strconn2)
						Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
						Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
						Strsql = Strsql & " where a.codobjmaster=" & rse(0)
						Strsql = Strsql & " and a.tipobjmaster='" & rse(2)
						Strsql = Strsql & "' And a.estado = 3 "

						If tipo <> "TODOS" Then
							Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
						End If

					'	If session("EMPMUN") = "S" Then
							Strsql = Strsql & " And a.objun =" & session("UN")
						'End If

						if desde = "TODOS"  then
									else
									Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
									end if


									if  hasta = "TODOS" then
									Strsql = Strsql & ""
									else
									 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
									end if


						Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
						Strsql = Strsql & "  "
						Set rsf = ofv.crearconsultaEx(StrSql,cnf,1,parametros)
						Do Until rsf.eof
							tipox = "Esquema: "
							For inx=0 To UBound(tobj)
								If tobj(inx) = rsf(2) Then
									tipox = tobjd(inx) & ": "
									Exit For
								End If
							Next
							subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
							subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
							subtemax = subtemax & "<td width=5>&nbsp;</td>"
							'subtemax = subtemax & "<td colspan=8  class=ut >" & tipox & rsf(1) & "</td>"
							If session("APLMDCLI")="BC" Then
								subtemax = subtemax & "<td colspan=8  class=ut >" & tipox & rsF("CODIDEN") & " : " & RSF(1) & "</td>"
							Else
								subtemax = subtemax & "<td colspan=8  class=ut >" & tipox & rsF(1) & "</td>"
							End If

							''  subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsf(4) + nFecini) & "</td>"

							dvige = CStr(rsf(4) & "")
							If CLng("0" & dvige) > 0 Then
								subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
							Else
								subtemax = subtemax & "<td class=ut >&nbsp;</td>"
							End If

							subtemax = subtemax & "</tr>"

							Call linea(subtemax,indi)
							'xx Arma Objetos G:
							Set cng = ofv.conectar(ofv.strconn2)
							Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
							Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
							Strsql = Strsql & " where a.codobjmaster=" & rsf(0)
							Strsql = Strsql & " and a.tipobjmaster='" & rsf(2)
							Strsql = Strsql & "' And a.estado = 3 "

							If tipo <> "TODOS" Then
								Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
							End If

						'	If session("EMPMUN") = "S" Then
								Strsql = Strsql & " And a.objun =" & session("UN")
						'	End If

						if desde = "TODOS"  then
									else
									Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
									end if


									if  hasta = "TODOS" then
									Strsql = Strsql & ""
									else
									 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
									end if


							Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
							Strsql = Strsql & "  "
							Set rsg = ofv.crearconsultaEx(StrSql,cng,1,parametros)
							Do Until rsg.eof
								tipox = "Esquema: "
								For inx=0 To UBound(tobj)
									If tobj(inx) = rsg(2) Then
										tipox = tobjd(inx) & ": "
										Exit For
									End If
								Next
								subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
								subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
								subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
								'subtemax = subtemax & "<td colspan=7  class=ut >" & tipox & rsg(1) & "</td>"
								If session("APLMDCLI")="BC" Then
									subtemax = subtemax & "<td colspan=7 class=ut >" & tipox & rsG("CODIDEN") & " : " & RSG(1) & "</td>"
								Else
									subtemax = subtemax & "<td colspan=7 class=ut >" & tipox & rsG(1) & "</td>"
								End If

								'''   subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsg(4) + nFecini) & "</td>"


								dvige = CStr(rsg(4) & "")
								If CLng("0" & dvige) > 0 Then
									subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
								Else
									subtemax = subtemax & "<td class=ut >&nbsp;</td>"
								End If


								subtemax = subtemax & "</tr>"
								Call linea(subtemax,indi)
								'xx Arma Objetos H:
								Set cnh = ofv.conectar(ofv.strconn2)
								Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
								Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
								Strsql = Strsql & " where a.codobjmaster=" & rsg(0)
								Strsql = Strsql & " and a.tipobjmaster='" & rsg(2)
								Strsql = Strsql & "' And a.estado = 3 "

								If tipo <> "TODOS" Then
									Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
								End If

							'	If session("EMPMUN") = "S" Then
									Strsql = Strsql & " And a.objun =" & session("UN")
							'	End If

							if desde = "TODOS"  then
										else
										Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
										end if


										if  hasta = "TODOS" then
										Strsql = Strsql & ""
										else
										 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
										end if


								Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
								Strsql = Strsql & "  "
								Set rsh = ofv.crearconsultaEx(StrSql,cnh,1,parametros)
								Do Until rsh.eof
									tipox = "Esquema: "
									For inx=0 To UBound(tobj)
										If tobj(inx) = rsh(2) Then
											tipox = tobjd(inx) & ": "
											Exit For
										End If
									Next
									subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
									subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
									subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
									subtemax = subtemax & "<td width=5>&nbsp;</td>"
									'subtemax = subtemax & "<td colspan=6  class=ut >" & tipox & rsh(1) & "</td>"
									If session("APLMDCLI")="BC" Then
										subtemax = subtemax & "<td colspan=6  class=ut >" & tipox & rsH("CODIDEN") & " : " & RSH(1) & "</td>"
									Else
										subtemax = subtemax & "<td colspan=6  class=ut >" & tipox & rsH(1) & "</td>"
									End If

									'''  subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsh(4) + nFecini) & "</td>"

									dvige = CStr(rsh(4) & "")
									If CLng("0" & dvige) > 0 Then
										subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
									Else
										subtemax = subtemax & "<td class=ut >&nbsp;</td>"
									End If


									subtemax = subtemax & "</tr>"
									Call linea(subtemax,indi)
									'xx Arma Objetos I:
									Set cni = ofv.conectar(ofv.strconn2)
									Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
									Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
									Strsql = Strsql & " where a.codobjmaster=" & rsh(0)
									Strsql = Strsql & " and a.tipobjmaster='" & rsh(2)
									Strsql = Strsql & "' And a.estado = 3 "

									If tipo <> "TODOS" Then
										Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
									End If

								'	If session("EMPMUN") = "S" Then
										Strsql = Strsql & " And a.objun =" & session("UN")
								'	End If

								if desde = "TODOS"  then
											else
											Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
											end if


											if  hasta = "TODOS" then
											Strsql = Strsql & ""
											else
											 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
											end if

									Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
									Strsql = Strsql & "  "
									Set rsi = ofv.crearconsultaEx(StrSql,cni,1,parametros)
									Do Until rsi.eof
										tipox = "Esquema: "
										For inx=0 To UBound(tobj)
											If tobj(inx) = rsi(2) Then
												tipox = tobjd(inx) & ": "
												Exit For
											End If
										Next
										subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
										subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
										subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
										subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
										'subtemax = subtemax & "<td colspan=5  class=ut >" & tipox & rsi(1) & "</td>"
										If session("APLMDCLI")="BC" Then
											subtemax = subtemax & "<td colspan=5  class=ut >" & tipox & rsI("CODIDEN") & " : " & RSI(1) & "</td>"
										Else
											subtemax = subtemax & "<td colspan=5  class=ut >" & tipox & rsI(1) & "</td>"
										End If

										'''   subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsi(4) + nFecini) & "</td>"

										dvige = CStr(rsi(4) & "")
										If CLng("0" & dvige) > 0 Then
											subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
										Else
											subtemax = subtemax & "<td class=ut >&nbsp;</td>"
										End If


										subtemax = subtemax & "</tr>"
										Call linea(subtemax,indi)
										'xx Arma Objetos J:
										Set cnj = ofv.conectar(ofv.strconn2)
										Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
										Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
										Strsql = Strsql & " where a.codobjmaster=" & rsi(0)
										Strsql = Strsql & " and a.tipobjmaster='" & rsi(2)
										Strsql = Strsql & "' And a.estado = 3 "

										If tipo <> "TODOS" Then
											Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
										End If

								'		If session("EMPMUN") = "S" Then
											Strsql = Strsql & " And a.objun =" & session("UN")
									'	End If

									if desde = "TODOS"  then
												else
												Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
												end if


												if  hasta = "TODOS" then
												Strsql = Strsql & ""
												else
												 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
												end if

										Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
										Strsql = Strsql & "  "
										Set rsj = ofv.crearconsultaEx(StrSql,cnj,1,parametros)
										Do Until rsj.eof
											tipox = "Esquema: "
											For inx=0 To UBound(tobj)
												If tobj(inx) = rsj(2) Then
													tipox = tobjd(inx) & ": "
													Exit For
												End If
											Next
											subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
											subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
											subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
											subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
											subtemax = subtemax & "<td width=5>&nbsp;</td>"
											'subtemax = subtemax & "<td colspan=4  class=ut >" & tipox & rsj(1) & "</td>"
											If session("APLMDCLI")="BC" Then
												subtemax = subtemax & "<td colspan=4  class=ut >" & tipox & rsJ("CODIDEN") & " : " & RSJ(1) & "</td>"
											Else
												subtemax = subtemax & "<td colspan=4 class=ut >" & tipox & rsJ(1) & "</td>"
											End If

											'''     subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsj(4) + nFecini) & "</td>"

											dvige = CStr(rsj(4) & "")
											If CLng("0" & dvige) > 0 Then
												subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
											Else
												subtemax = subtemax & "<td class=ut >&nbsp;</td>"
											End If


											subtemax = subtemax & "</tr>"
											Call linea(subtemax,indi)
											'xx Arma Objetos K:
											Set cnk = ofv.conectar(ofv.strconn2)
											Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
											Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
											Strsql = Strsql & " where a.codobjmaster=" & rsj(0)
											Strsql = Strsql & " and a.tipobjmaster='" & rsj(2)
											Strsql = Strsql & "' And a.estado = 3 "

											If tipo <> "TODOS" Then
												Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
											End If

										'	If session("EMPMUN") = "S" Then
												Strsql = Strsql & " And a.objun =" & session("UN")
										'	End If
if desde = "TODOS"  then
			else
			Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
			end if


			if  hasta = "TODOS" then
			Strsql = Strsql & ""
			else
			 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
			end if

											Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
											Strsql = Strsql & "  "

											Set rsk = ofv.crearconsultaEx(StrSql,cnk,1,parametros)
											Do Until rsk.eof
												tipox = "Esquema: "
												For inx=0 To UBound(tobj)
													If tobj(inx) = rsk(2) Then
														tipox = tobjd(inx) & ": "
														Exit For
													End If
												Next
												subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
												subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
												subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
												subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
												subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
												' subtemax = subtemax & "<td colspan=3  class=ut >" & tipox & rsk(1) & "</td>"
												If session("APLMDCLI")="BC" Then
													subtemax = subtemax & "<td colspan=3  class=ut >" & tipox & rsK("CODIDEN") & " : " & RSK(1) & "</td>"
												Else
													subtemax = subtemax & "<td colspan=3  class=ut >" & tipox & rsK(1) & "</td>"
												End If

												''   subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsk(4) + nFecini) & "</td>"

												dvige = CStr(rsk(4) & "")
												If CLng("0" & dvige) > 0 Then
													subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
												Else
													subtemax = subtemax & "<td class=ut >&nbsp;</td>"
												End If

												subtemax = subtemax & "</tr>"
												Call linea(subtemax,indi)
												'xx Arma Objetos L:
												Set cnl = ofv.conectar(ofv.strconn2)
												Strsql = "SELECT distinct a.*, b.nroacta, b.fecacta from csLisObjetos A "
												Strsql = Strsql & " inner join objprty b on (a.codobjeto = b.codobjeto) "
												Strsql = Strsql & " where a.codobjmaster=" & rsk(0)
												Strsql = Strsql & " and a.tipobjmaster='" & rsk(2)
												Strsql = Strsql & "' And a.estado = 3 "

												If tipo <> "TODOS" Then
													Strsql = Strsql & " and a.tipobjeto IN " & SQLTIPO
												End If

											'	If session("EMPMUN") = "S" Then
													Strsql = Strsql & " And a.objun =" & session("UN")
											'	End If

												if desde = "TODOS"  then
															else
															Strsql = Strsql & " And a.fecvigen >  '" & (desde - 1) & "' "
															end if


															if  hasta = "TODOS" then
															Strsql = Strsql & ""
															else
															 Strsql = Strsql & "    and a.fecvigen <  '" & (hasta + 1) & "'  "
															end if
												Strsql = Strsql & " order by a.tipobjeto, a.descri, a.codobjeto"
												Strsql = Strsql & "  "

												Set rsl = ofv.crearconsultaEx(StrSql,cnl,1,parametros)
												Do Until rsl.eof
													tipox = "Esquema: "
													For inx=0 To UBound(tobj)
														If tobj(inx) = rsl(2) Then
															tipox = tobjd(inx) & ": "
															Exit For
														End If
													Next
													subtemax = "<tr><td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
													subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
													subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
													subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
													subtemax = subtemax & "<td width=5>&nbsp;</td><td width=5>&nbsp;</td>"
													subtemax = subtemax & "<td width=5>&nbsp;</td>"
													'subtemax = subtemax & "<td colspan=2  class=ut >" & tipox & rsl(1) & "</td>"`
													If session("APLMDCLI")="BC" Then
														subtemax = subtemax & "<td colspan=2 class=ut >" & tipox & rsL("CODIDEN") & " : " & RSL(1) & "</td>"
													Else
														subtemax = subtemax & "<td colspan=2  class=ut >" & tipox & rsL(1) & "</td>"
													End If


													'''    subtemax = subtemax & "<td class=ut >Vigencia: " & cdate(rsl(4) + nFecini) & "</td>"

													dvige = CStr(rsl(4) & "")
													If CLng("0" & dvige) > 0 Then
														subtemax = subtemax & "<td class=ut >Vigencia: " & fechadorCor(CDate(dvige + nFecini),"") & "</td>"
													Else
														subtemax = subtemax & "<td class=ut >&nbsp;</td>"
													End If


													subtemax = subtemax & "</tr>"
													Call linea(subtemax,indi)

													rsl.movenext
												Loop
												Call ofv.cerrarconsulta(rsl)
												Call ofv.cerrarconn(cnl)

												rsk.movenext
											Loop
											Call ofv.cerrarconsulta(rsk)
											Call ofv.cerrarconn(cnk)

											rsj.movenext
										Loop
										Call ofv.cerrarconsulta(rsj)
										Call ofv.cerrarconn(cnj)

										rsi.movenext
									Loop
									Call ofv.cerrarconsulta(rsi)
									Call ofv.cerrarconn(cni)


									rsh.movenext
								Loop
								Call ofv.cerrarconsulta(rsh)
								Call ofv.cerrarconn(cnh)

								rsg.movenext
							Loop
							Call ofv.cerrarconsulta(rsg)
							Call ofv.cerrarconn(cng)

							rsf.movenext
						Loop
						Call ofv.cerrarconsulta(rsf)
						Call ofv.cerrarconn(cnf)

						rse.movenext
					Loop
					Call ofv.cerrarconsulta(rse)
					Call ofv.cerrarconn(cne)

					rsd.movenext
				Loop
				Call ofv.cerrarconsulta(rsd)
				Call ofv.cerrarconn(cne)

				rsc.movenext
			Loop
			Call ofv.cerrarconsulta(rsc)
			Call ofv.cerrarconn(cnc)

			rsb.movenext
		Loop
		Call ofv.cerrarconsulta(rsb)
		rsa.movenext
	Loop
	Call ofv.cerrarconsulta(rsa)
	Call ofv.cerrarconn(cn)

	response.write "</table>"
	response.write "</td>"
	response.write "</tr>"
	response.write "</table>"


	Agregar = ""
	Buscar  = ""
	Celda   = ""

	response.write  ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


	response.write "</body></html>"
End If

response.write sHTML

%>
