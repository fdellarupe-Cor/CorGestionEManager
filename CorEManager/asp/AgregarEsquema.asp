<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo = request.querystring("Tipo")
sNombre = request.querystring("Nombre")
sUsuario = session("Usuario")
sElemento = request.querystring("Elemento")
sACC = request.querystring("ACC")
idpedido=0
'sACC  identifica si viene de pedidos
If sACC <> "" Then

	sTipo = request.form("Tipo")
	sNombre = request.form("Nombre")
	sElemento = request.form("Elemento")
	idpedido=request.querystring("idpedido")
End If
n = InStr(sElemento,"-")
If n > 0 Then
	nCodMaster = Left(sElemento,n-1)
	sTipoMaster = Mid(sElemento,n+1)
	n = InStr(sTipoMaster,"-")
	sTipoMaster = Left(sTipoMaster,n-2)
End If

SPTIPO = request.querystring("SPTIPO")

'response.write " elemento " & selemento & " tipomaster " & stipomaster

On Error Resume Next

Set cn = ofv.conectar(ofv.strconn2)
nFecini = 0



If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

strsql = "select * from Objetos"
Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert into Objetos (Tipobjeto"
If sacc <> "" Then

	For Each x In Request.form
		x=LCase(x)
		If x <> "tipo" And x <> "nombre" And x <> "elemento" Then
			strsql = strsql & "," & x
		End If
	Next
Else
	For Each x In Request.form
		strsql = strsql & "," & x
	Next
End If


strsql = strsql & ") Values ('" & sTipo & "'"
If sacc <> "" Then

	For Each yy In Request.form
		x = LCase(yy)
		If x <> "nombre" And x <> "tipo" And x <> "elemento" Then
			If x = "aplic" Then WKFAPLIC = Request.Form(x)
			TipoValido = ofv.ValidarTipo(rs(x),Request.Form(x))
			If TipoValido = False And LCase(rs(x).name) <> "codiden" Then Exit For
			ValorNulo = ofv.EsNulo(rs(x),Request.Form(x))
			If ValorNulo = True And rs(x).name <> "codiden" Then Exit For
			s = ofv.queesx(rs(x).type)
			If rs(x).type = 200 Then s = "'"
			If LCase(rs(x).name) = "codiden" Then
				strsql = strsql & ",'" & CStr(session.sessionid) & "'"
				TipoValido = True
				ValorNulo = False
			Else
				strsql = strsql & "," & s & UCase(Request.Form(x)) & s
			End If
		End If
	Next
Else
	For Each yy In Request.form

		x = LCase(yy)



		If x = "aplic" Then WKFAPLIC = Request.Form(x)
		TipoValido = ofv.ValidarTipo(rs(x),Request.Form(x))
		If TipoValido = False And LCase(rs(x).name) <> "codiden" Then Exit For
		ValorNulo = ofv.EsNulo(rs(x),Request.Form(x))
		If ValorNulo = True And rs(x).name <> "codiden" Then Exit For
		s = ofv.queesx(rs(x).type)
		If rs(x).type = 200 Then s = "'"
		If LCase(rs(x).name) = "codiden" Then
			strsql = strsql & ",'" & CStr(session.sessionid) & "'"
			TipoValido = True
			ValorNulo = False
		Else
			strsql = strsql & "," & s & UCase(Request.Form(x)) & s
		End If
	Next

End If


volver = "../vuelve.asp"
If TipoValido = False Or ValorNulo = True Then
	On Error Resume Next
	response.write ofv.ErrorTipo(volver,Request.Form(x),rs(nx).name)
	If Err.number <> 0 Then response.write ofv.ErrorTipo(volver,"","")
Else
	strsql = strsql & ")"
	' response.write "primer insert " & strsql & "<br>"
	Call ofv.crearconsultaEx(StrSql,cn,1,volver)

	If cn.errors.count = 0 Then
		Call ofv.cerrarconsulta(rs)

		strsql = "Select codobjeto From Objetos Where codiden = '" & CStr(session.sessionid) & "'"

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then nCodAsoc = rs(0)

		Call ofv.cerrarconsulta(rs)

		If request.form("codiden") = "" Then
			ncodiden = "SD"
		Else
			ncodiden =  request.form("codiden")
		End If

		strsql = "Update Objetos set codiden = '" & ncodiden & "', indobj = 10 "
		strsql = strsql & " where codiden = '" & CStr(session.sessionid) & "'"
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		Strsql = "SELECT * FROM OBJdistPrty  "
		Strsql = Strsql & " Where codobjeto=0"
		Strsql = Strsql & " And tipobjeto = 'DEFDST' "

		Set OBJdistPrty = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		If OBJdistPrty.eof Then
			Strsql = "insert into OBJdistPrty  (codobjeto,tipobjeto,emisor,asunto,texto,importancia,aviso,dias) values ("
			Strsql = Strsql & nCodAsoc & ",'" & stipo & "','" & session("usuario") & "','Activacion de <ESQNOR>',"
			Strsql = Strsql & "'<ESQNOR> esta vigente a partir de la fecha',1,'NO','7')"
		Else
			Strsql = "insert into OBJdistPrty  (codobjeto,tipobjeto,emisor,asunto,texto,importancia,aviso,dias) values ("
			Strsql = Strsql & nCodAsoc & ",'" & stipo & "','COREMANAGER','"& OBJdistPrty("asunto")& "',"

			Strsql = Strsql & "'" & OBJdistPrty("texto")&"',"& OBJdistPrty("importancia") &",'"& OBJdistPrty("aviso")&"','"& OBJdistPrty("dias")&"')"
		End If
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)
		' response.write "2do insert " & strsql & "<br>"
		Call ofv.cerrarconsulta(OBJdistPrty)
		sNIVE = "D"

		StrSql = "Select TIPOCPT From Tipobjeto "
		StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & stipo & "'"
		Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
		If  Not rsIcono.EOF Then
			sNIVE  = rsIcono(0)
		End If
		Call OFV.cerrarconsulta(rsIcono)


		If SESSION("EMPMUN") = "S" Then

			Strsql = "insert into OBJdist  (codobjeto,tipobjeto,GRPUSR,SECTOR) values ("
			Strsql = Strsql & nCodAsoc & ",'" & stipo & "','TODOS','" & session("UN") & "')"
			Call ofv.crearconsultaEx(StrSql,cn,1,volver)
			'    response.write "3 insert " & strsql & "<br>"
		End If

		strsql = "Insert into Relaobj (Codobjmaster,Tipobjmaster,Codojasoc,Tipobjasoc) "
		strsql = strsql & "Values (" & nCodMaster & ",'" & sTipomaster & "',"
		strsql = strsql & nCodAsoc & ",'" & sTipo & "')"
		'    response.write "4 insert " & strsql & "<br>"
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)
		nEstado = 1
		If cn.errors.count = 0 Then

			If sNIVE = "D" Or sNIVE = "K" Then
				strsql = "Select objid From Estados Where estado = '" & UCase(session("menu")) & "'"

				' response.write "estado " & strsql & "<br>"
				Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
				If Not rs.eof Then nEstado = rs(0)
			Else
				nEstado = 3
			End If

			If WKFAPLIC = "RG" Then nEstado = 3
			If WKFAPLIC = "PL" Then nEstado = 1

			'      dFechaNula = "1/1/1900"
			dFechaNula = CLng(CDate("01-01-1900")) - nFecini
			dHoy = CLng(Date) - nFecini
			strsql = "Insert into Objprty (codobjeto, tipobjeto, estado, fecemis, fecvigen, "
			strsql = strsql & "fecbaja, fecvto, fecacta,fecpublic, autor) Values (" & nCodAsoc & ",'" & sTipo & "',"
			strsql = strsql & nEstado & ",'" & DHoy & "','" & DFechaNula & "','"
			strsql = strsql & DFechaNula & "','" & DFechaNula & "','" & dFechaNula & "','" & dFechaNula & "','" & sUsuario & "')"

			'  response.write "5 insert " & strsql & "<br>"

			Call ofv.crearconsultaEx(StrSql,cn,1,volver)
			If cn.errors.count = 0 Then
				'response.write sNIVE
				If (sNIVE="K") And  (session("APLMDCLI")="BGAL"  OR session("CORMODUSOLABM")="S") Then

					If sacc <> "" Then 'entra si es un pedido


						strsql= "select a.pos,a.descripcion,b.objun from pedidosarch a "
						strsql = strsql &  " inner join pedidos b on(b.objid=a.idpedido) "
						strsql =strsql & " where a.idpedido=" &  idpedido
					'	response.write strsql
						Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
						If SESSION("FILETIPO")="L" Then
							Set fso = Server.CreateObject("scripting.filesystemobject")
							sPath = Server.mappath(Request.servervariables("SCRIPT_NAME"))
							sPath = Mid(sPath, 1, InStrRev(sPath, "\", -1)-1)
							sPath = Mid(sPath, 1, InStrRev(sPath, "\", -1))
							sPatH =session ("DirUpload")
							spath=Replace(UCase(spath),"\ASP\","\")
						End If

                        pedobjun=1
						Do While Not rsIcono.eof
						    pedobjun=rsIcono("objun")
						    if Session("EMPMUN")  = "S" then
						     fso.CopyFile sPatH  & "emp" & rsIcono("objun") & "\pedidos\" &  rsIcono("descripcion"), sPatH  & "emp" & rsIcono("objun") & "\ARCCOMPONENTES\" &  rsIcono("descripcion")
							 sql3 = "insert into objhtm (codobjeto,tipobjeto,secuen,descri,archivo,tipo,directorio,modi,tipmodi,modautor,modfecha) "
							 sql3 = sql3 & " values( "  & nCodAsoc & ",'" & sTipo & "'," & rsIcono("pos") & ",'','" & rsIcono("descripcion") & "','1','ARCCOMPONENTES',1,'M','"& session("usuario") & "'," & dHoy & ")"

						    else
							 fso.CopyFile sPatH & "pedidos\"&  rsIcono("descripcion"), sPatH & "ARCCOMPONENTES\" &  rsIcono("descripcion")
							 sql3 = "insert into objhtm (codobjeto,tipobjeto,secuen,descri,archivo,tipo,directorio,modi,tipmodi,modautor,modfecha) "
							 sql3 = sql3 & " values( "  & nCodAsoc & ",'" & sTipo & "'," & rsIcono("pos") & ",'','" & rsIcono("descripcion") & "','1','ARCCOMPONENTES',1,'M','"& session("usuario") & "'," & dHoy & ")"
							end if
							Call OFV.crearconsultaEx(sql3, cn, 1, "")
							rsIcono.movenext
						Loop
						Call ofv.cerrarconsulta(rsIcono)



						strsql= "select * from pedidosdet where idpedido=" &  idpedido
						'response.write strsql


						Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")


						If Not rsIcono.eof Then
							strsql = "Insert into OBJPROPIEDADES (codobjeto,tipobjeto"
							For ii = 2 To 12
								If Not IsNull(rsIcono(ii))  Then
									strsql = strsql & "," & rsIcono(ii).Name
								End If
							Next

							If Not IsNull(rsIcono("producto")) Then
								strsql = strsql & "," & "producto"
							End If


							strsql = strsql & ") values (" & nCodAsoc & ",'" & sTipo & "'"



							strsql2 = ""

							For ii = 2 To 12
								If Not IsNull(rsIcono(ii)) Then
									sInput = rsIcono(ii)
									strsql2 = strsql2 & ",'" & sInput & "'"
								End If
							Next

							If Not IsNull(rsIcono("producto")) Then
								sInput = rsIcono("producto")
								strsql2 = strsql2 & ",'" & sInput & "'"

								sqlz="update objetos set tsegmento="	& sInput & " where codobjeto=" & 	nCodAsoc
								Call OFV.crearconsultaEx(sqlz, cn, 1, volver)

							End If

							Call OFV.CERRARCONSULTA(rsIcono)
							'response.write strsql & strsql2 & ") "
							Call OFV.crearconsultaEx(strsql & strsql2 & ") ", cn, 1, volver)

							SQL = "Select ARCHIVO,A.TIPOBJETO from  pedidos A "
							SQL = SQL & "  INNER JOIN TIPOBJETO B ON (A.TIPOBJETO=B.TIPOBJETO)"
							SQL= SQL & " where A.OBJID=0" &  idpedido
							'RESPONSE.WRITE SQL

							Set rsIcono = OFV.crearconsultaEx(SQL, cn, 1, "")


							If Not rsIcono.eof Then
								DESTINO=Split(rsIcono("ARCHIVO"),"-")
								' RESPONSE.WRITE rsIcono("ARCHIVO")
								If IsArray(DESTINO) Then
									If  UBound(DESTINO)=1 Then
										nSUBTEMA=DESTINO(1)
										STIPOSUBTEMA=DESTINO(0)

										PEDIDOTIPO=rsIcono(1)
										SQL="UPDATE RELAOBJ SET TIPOBJMASTER='" & STIPOSUBTEMA & "',CODOBJMASTER=" & nSUBTEMA & "  WHERE CODOJASOC=" & idpedido  & " AND TIPOBJASOC='" & PEDIDOTIPO & "'"
										'      RESPONSE.WRITE "<BR>" & SQL
										Call OFV.crearconsultaEx(sql, cn, 1, volver)
									End If
								End If
							End If

							Call OFV.CERRARCONSULTA(rsIcono)

							sql="update pedidos set estado=5  where objid=" &  idpedido


							Call OFV.crearconsultaEx(sql, cn, 1, volver)
						Else
							Call OFV.cerrarconsulta(rsIcono)
						End If



					Else
						sql=" insert into objpropiedades (codobjeto,tipobjeto) values (" & nCodAsoc & " ,'" & sTipo &  "')"


					End If

					Call ofv.crearconsultaEx(SQL,cn,1,volver)
					sqlz= "update contador set codigof=codigof + 1 , codigov= (codigov + 1 ) % 10"

					Call ofv.crearconsultaEx(SQLz,cn,1,volver)


				End If
				strsql = "Insert into OBJSDK (codobjeto, tipobjeto, segct,segmo,segco,segce,segown "
				strsql = strsql & ") Values (" & nCodAsoc & ",'" & sTipo & "',0,0,0,0,'"
				strsql = strsql & Session("Usuario") & "')"
				'     response.write "6 insert " & strsql & "<br>"
				Call ofv.crearconsultaEx(StrSql,cn,1,volver)

				' SE AGREGA ALCANCE 22/9
                 if sacc <> "" then ' verifica que es un pedido
                  strsql = "Insert into OBJMEMPALC (codobjeto, tipobjeto, codempre"
				  strsql = strsql & ") Values (" & nCodAsoc & ",'" & sTipo & "','" & pedobjun & "')"
				  Call ofv.crearconsultaEx(StrSql,cn,1,volver)
                 ' strsql = "Insert into OBJMEMPALC (codobjeto, tipobjeto, codempre"
				'  strsql = strsql & ") Values (" & nCodAsoc & ",'" & sTipo & "','" & pedobjun & "')"
				' Call ofv.crearconsultaEx(StrSql,cn,1,volver)
                 else
					If session("APLMDCLI")="GP" Then
					 strsql = "Insert into OBJMEMPALC (codobjeto, tipobjeto, codempre"
					 strsql = strsql & ") Values (" & nCodAsoc & ",'" & sTipo & "','" & SESSION("CLIEMP") & "')"
					 Call ofv.crearconsultaEx(StrSql,cn,1,volver)
					End If
				end if

				If cn.errors.count = 0 Then
					strsql = "Insert into OBJWFL (codobjeto, tipobjeto, fecvto)"
					strsql = strsql & " Values (" & nCodAsoc & ",'" & sTipo & "','1')"
					'      response.write "7 insert " & strsql & "<br>"
					Call ofv.crearconsultaEx(StrSql,cn,1,volver)
					If cn.errors.count = 0 Then

						'''''''''''' SOLO PARA SUBPROCESOS CON VINCULACION EXPLICITA (ACTIVIDADES ADMINISTRADAS EN FORMA MANUAL)
						If SPTIPO = "CR" Or SPTIPO = "AI" Then


							GETFECH()

							WKFECHA = XLFECDAT - CLng(CDate("01-01-1900"))
							WKHORA  = XLHORDAT
							WKSECTOR = 0

							strsql = "insert into OBJSECTOR  ("
							strsql = strsql & "codobjeto,tipobjeto,codSECTOR,MODI,"
							strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA"
							strsql = strsql & ") values ("
							strsql = strsql & nCodAsoc & ",'" & sTipo & "','0','1',"
							strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "'"
							strsql = strsql & ")"
							'        response.write "8 insert " & strsql & "<br>"
							Call ofv.crearconsultaEx(StrSql,cn,1,volver)

							strsql = "SELECT OBJID FROM OBJSECTOR "
							strsql = strsql & " WHERE CODOBJETO = " & nCodAsoc & " "
							Set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If Not rsN.eof Then
								WKSECTOR = rsN(0)
							End If


							sevento = 0
							sukey = GTkey(nCodAsoc)

							strsql = "insert into objproced  ("
							strsql = strsql & "codobjeto,tipobjeto,codproced,codeven,"
							strsql = strsql & "frecuencia,tipproced,CODSECTOR,MODI,"
							strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA"
							strsql = strsql & ") values ("
							strsql = strsql & nCodAsoc & ",'" & sTipo & "','" & sukey & "','" & sevento & "',"
							strsql = strsql & "'DIARIA','" & wkaplic & "','" & WKSECTOR & "','1',"
							strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "'"
							strsql = strsql & ")"
							'          response.write "9 insert " & strsql & "<br>"
							Call ofv.crearconsultaEx(StrSql,cn,1,volver)




							strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
							strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values ('" & nCodMaster & "','" & sTipomaster & "',"
							strsql = strsql & "'" & nCodMaster & "','" & nCodMaster & "',"
							strsql = strsql & "'" & nCodAsoc & "','" & sTipo & "','','','LRGCR','5','" & sukey & "','0'"
							strsql = strsql & ")"
							'    response.write "10 insert " & strsql & "<br>"
							Call ofv.crearconsultaEx(StrSql,cn,1,volver)

							Strsql = "SELECT * From OBJLNK where iden1 = '" & nCodMaster & "' and iden2 = '" & nCodMaster & "' and LTIPO = 'LRGCR' "
							Set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If Not rsN.eof Then
								wlnkid = rsN(0)
							Else
								wlnkid = 0
								wlink = 0
							End If
							Call ofv.cerrarconsulta(rsN)

							If CLng(wlnkid) > 0 Then
								strsql = "update OBJETOS set rellink = " & wlnkid
								strsql = strsql & " where  codobjeto = 0" & nCodMaster
								Call ofv.crearconsultaEx(StrSql,cn,1,volver)
							End If



						End If

						'''''''''''' SOLO PARA OBJETOS SOPORTE DEL PLAN DE CONTINGENCIAS
						If SPTIPO = "PL" Then


							GETFECH()

							WKFECHA = XLFECDAT - CLng(CDate("01-01-1900"))
							WKHORA  = XLHORDAT
							WKSECTOR = 0


							strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, iden3, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
							strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values ('" & nCodMaster & "','" & sTipomaster & "',"
							strsql = strsql & "'" & nCodMaster & "','" & nCodMaster & "','" & SESSION("ORPERIODO") & "',"
							strsql = strsql & "'" & nCodAsoc & "','" & sTipo & "','','','LRGPL','0','0','0'"
							strsql = strsql & ")"
							'          response.write "11 insert " & strsql & "<br>"
							Call ofv.crearconsultaEx(StrSql,cn,1,volver)

							Strsql = "SELECT * From OBJLNK where iden1 = '" & nCodMaster & "' and "
							strsql = strsql & " iden2 = '" & nCodMaster & "' and "
							strsql = strsql & " iden3 = '" & SESSION("ORPERIODO") & "' and "
							strsql = strsql & "LTIPO = 'LRGPL' "
							Set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If Not rsN.eof Then
								wlnkid = rsN(0)
							Else
								wlnkid = 0
								wlink = 0
							End If
							Call ofv.cerrarconsulta(rsN)

							If CLng(wlnkid) > 0 Then
								strsql = "update OBJETOS set rellink = " & wlnkid
								strsql = strsql & " where  codobjeto = 0" & nCodMaster
								Call ofv.crearconsultaEx(StrSql,cn,1,volver)
							End If



						End If

						'SE ACLARO QUE LA DESCRI SE PODIA MODIFICAR EN UN FUTURO Y SE DESPAREJEN LOS CODIGOFORM, KARINA ACEPTO
						If (sNIVE="K" And  session("APLMDCLI")="BGAL")  Then

							ssql = "select codigoform from objpropiedades where codobjeto=" & nCodAsoc & " and tipobjeto='" & sTipo & "'"
							Set s2 = ofv.crearconsultaEx(ssql,cn,1,volver)
							cf = ""
							If Not s2.eof Then cf = s2(0)

							xdsc = "F" & cf & " "

							ssql = "update objetos set descri='" & xdsc & "' + descri where codobjeto=" & nCodAsoc & " and tipobjeto='" & sTipo & "'"
							Call ofv.crearconsultaEx(ssql,cn,1,volver)


							Call ofv.cerrarconsulta(s2)

						End If

						response.redirect volver
					Else
						response.write ofv.ErrorDB(cn,volver)
					End If
				Else
					response.write ofv.ErrorDB(cn,volver)
				End If
			Else
				response.write ofv.ErrorDB(cn,volver)
			End If
		Else
			response.write ofv.ErrorDB(cn,volver)
		End If
	Else
		response.write ofv.ErrorDB(cn,volver)
	End If
End If

Call ofv.cerrarconsulta(rs)
Call ofv.cerrarconn(cn)

%>