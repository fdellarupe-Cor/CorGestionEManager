<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok=ofv.CheckUsuario()

If ok <> "Y" Then
	sHTML = ok
Else
	response.expires=0


	sHTML = ofv.MenuHeader("","")



	'----------------------------------------------

	shtml = shtml & CorPnlFnc("N","")

	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

	'-------------------------------------------------

	CantRegMostrarx = 12
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

	Set cn = ofv.conectar(ofv.strconn5)
	Set cn4 = ofv.conectar(ofv.strconn6)
	Set cn3 = ofv.conectar(ofv.strconn2)

	ofv.ObtenerAtributos session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar
	'response.write session("PASOOBS")
	If session("PASOOBS")="OBS" Then
		''' RESPONSE.WRITE "222" & session("PASOOBS") & "222222222222222"
		WKFAUTOB = "WKF00016631121722600921928999"
		WKFAPROB = "WKF00016631121722600921928999"
		WKFAUTO  = "WKF00016631121722600921928999"
		WKFAPRO  = "WKF00016631121722600921928999"
	Else
		WKFAUTOB = "WKF00016631121722600921928"
		WKFAPROB = "WKF00016631121722600921928"
		WKFAUTO  = "WKF00016631121722600921928"
		WKFAPRO  = "WKF00016631121722600921928"
	End If

	DESLEVL = 9
	SUPLEVL = 9


	If session("NCliente") <> "NO" Then

		SESSION("NOMREQ")=""

		SQL= "SELECT   'Solicitud nro:' + cast (a.objid as nvarchar)  + ': ' + BA.NOMBRE FROM PEDIDOS A "
		SQL= SQL  & "  INNER JOIN PEDIDOSDET BA ON(BA.IDPEDIDO=A.OBJID)"
		SQL= SQL  & "  INNER JOIN TIPOBJETO B ON(B.TIPOBJETO=A.TIPOBJETO)"
		SQL= SQL  & " WHERE A.OBJID=0" & session("NCliente") & " and a.tipobjeto='" &  session("TCliente") & "'"
		'RESPONSE.WRITE SQL
		'TIPOBJETO WHERE TIPOBJETO='" & TIPOBJETO & "'"
		Set rsx1 = ofv.crearconsultaEx(SQL,cn3,1,parametros)
		If Not rsx1.eof Then
			SESSION("NOMREQ")=UCase(rsx1(0) & "")
		End If

		Call ofv.cerrarconsulta(rsx1)

		If  SESSION("NOMREQ") <> ""  Then
			Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
			Set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If Not rs.eof Then


				WKFAUTO = RS(19)
				WKFAPRO = RS(20)
				tipocpt = rs(21)
				DESLEVL = 9
				SUPLEVL = RS(18)
			End If



		ElseIf session("TCliente") = "RSK" Then

			Strsql = "Select * From ORRSKCONFIG where APLIC = '" & WKAPLIC & "' "
			Set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If Not rs.eof Then

				WKFAUTO = RS(5)
				WKFAPRO = RS(6)

				DESLEVL = 9
				SUPLEVL = 8
			End If


			mstdsc = ""


			Strsql = "SELECT MSGCHAT FROM OBJWFL  "
			Strsql = Strsql & " Where codobjeto=" & session("NCliente")
			Strsql = Strsql & " And tipobjeto = '" & session("TCliente") & "'"
			Set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)

			If Not rs.eof Then
				mstdsc = RS(0)
			End If

			strsql = "SELECT * FROM  ORRSKCIRCUITO "
			strsql = strsql & " Where OBJID = 0" & session("NCliente")
			Set xrs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If Not xrs.eof Then
				If XRS("APLIC") = "OR" Then
					TIPOOBJ = "CIRCUITO AUTOEVALUACION"
					XSPROCESO=GetProceso(XRS(0))
				Else
					TIPOOBJ = "PLAN DE CONTINGENCIAS"
				End If

				Strsql = "Select * From objetos where codobjeto = 0" & XRS("CODOBJETO")
				Set xrs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
				If Not xrs.eof Then
					TIPOOBJ2 = XRS(1)
					TIPOOBJ3 = XRS(3)

				End If

				'SESSION("NOMREQ") = TIPOOBJ & ":&nbsp;" & TIPOOBJ2  & "&nbsp;-&nbsp;" & TIPOOBJ3
				If WKAPLIC="OR" Then
					SESSION("NOMREQ")=TIPOOBJ & ":&nbsp;" & XSPROCESO   '& "&nbsp;-&nbsp;" & TIPOOBJ3
				Else
					SESSION("NOMREQ") = TIPOOBJ & ":&nbsp;" & TIPOOBJ2  & "&nbsp;-&nbsp;" & TIPOOBJ3
				End If


			End If
			Call ofv.cerrarconsulta(xrs)

		Else


			Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
			Set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If Not rs.eof Then


				WKFAUTO = RS(19)
				WKFAPRO = RS(20)
				tipocpt = rs(21)
				DESLEVL = 9
				SUPLEVL = RS(18)
			End If

			mstdsc = ""


			Strsql = "SELECT MSGCHAT FROM OBJWFL  "
			Strsql = Strsql & " Where codobjeto=" & session("NCliente")
			Strsql = Strsql & " And tipobjeto = '" & session("TCliente") & "'"
			Set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)

			If Not rs.eof Then
				mstdsc = RS(0)
			End If


			'StrSql = "Select a.*,b.codigoform From OBJETOS a "
			If session ("base") ="K" Then
				StrSql = "Select a.*,b.codigoform From OBJETOS a "
				StrSql = strsql & " inner join objpropiedades b on (a.codobjeto = b.codobjeto and a.tipobjeto=b.tipobjeto)"
			Else
				StrSql = "Select a.* From OBJETOS a"
			End If
			StrSql = Strsql & " where a.CODOBJETO = " & session("NCliente")
			Set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If Not rs.eof Then

				Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
				Set rsX1 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
				If Not rsX1.eof Then
					TIPOOBJ = RSX1(1)
				End If
				Call ofv.cerrarconsulta(rsx1)


				If session ("base") ="K" Then
					SESSION("NOMREQ") = TIPOOBJ & ":&nbsp;" & RS(1) & "&nbsp;-&nbsp;" & RS(3)

				Else

					SESSION("NOMREQ") = TIPOOBJ & ":&nbsp;" & RS(1)  & "&nbsp;-&nbsp;" & RS(3)
				End If

			End If




		End If



	End If











	If (WKFAUTO & "") = "" Then WKFAUTO = WKFAUTOB
	If (WKFAPRO & "") = "" Then WKFAPRO = WKFAPROB

	sMODAL = request.querystring("MODAL")

	sopcion = request.querystring("opcion")
	sACCION = request.querystring("ACCION")
	act = request.querystring("act")
	SACC = request.querystring("acc")
	nCodigo = request.querystring("Codobjeto")
	sTipo   = request.querystring("Tipobjeto")
	'RESPONSE.WRITE SACC




	If nFecini < CLng(CDate("01-01-1900")) Then
		nFecini = CLng(CDate("01-01-1900"))
	Else
		nFecini = 0
	End If

	vNuevoValor = CLng(Date()) - nFecini

	''RESPONSE.WRITE SOPCION & " " & SACCION



	Select Case SOPCION

		Case "AUTO"
		SESSION("WKF") = WKFAUTO
		TITULO = SESSION("WKF") & " INICIO DE CIRCUITO DE AUTORIZACION "



		Case "APRO"
		SESSION("WKF") = WKFAPRO
		TITULO = SESSION("WKF") & " INICIO DE CIRCUITO DE OBSERVACIONES "


	End Select

	Dim ATABLA(200,6)

	Dim ATABSW(200,3)

	TSHW = 0




	''' GENERA EL PASO INICIAL

	StrSql = "Select * From WKFCKTMASTER WHERE INI=1 AND WKFCOD = '" & SESSION("WKF") & "' "
	''RESPONSE.WRITE STRSQL & "<BR>" & cn & "<BR>"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then

		'''  INICIA EL CIRCUITO

		session("WKFDSC") = rs(1)
		session("WKFCLASE") = rs(3)

		XMODEL  = RS(5)

		XRTIPO  = RS(15)
		XRCOD   = RS(16)



		XPASCOD = RS(12)
		XPDSC   = RS(6) & ": " & RS(7)
		XPDSC1   = RS(6) & ": " & RS(7)


		''' VER SI CONFIRMA ENVIO

		If act = "S" Then

			Set cn2 = ofv.conectar(ofv.strconn0)

			sInputx = "COR@MANAGER.COM"

			If SESSION("WKFETPU") = "D" Then               ' EMISOR DEFAULT (EL RESPONSABLE DEL PASO)

				xxxx = getInfoUser(SESSION("USUARIO"))
				If session("SGU_mail") <> "" Then    SINPUTX = SESSION("SGU_mail")

				SINPUTX =  SINPUTX & ""

			ElseIf SESSION("WKFETPU") = "U" Then           ' EMISOR PREDETERMINADO (PARAMETRIZADO EL LA BASE DE USUARIOS)

				WKFEUSR = SESSION("WKFEUSR")

				If WKFEUSR = "" Then   WKFEUSR = "COREMANAGER"


				strsql = "Select EMAIL From USUARIOS where ID = '" & WKFEUSR & "' "
				Set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
				If Not rsniv.eof Then sInputx = rsniv(0)
				Call ofv.cerrarconsulta(rsniv)

			Else                                           ' EMISOR FIJO (USUARIO Y MAIL FIJOS EN CORCONECT)

				sInputx = SESSION("WKFEMAI")


			End If



			If SINPUTX = "" Then SINPUTX = "COR@MANAGER.COM"

			SESSION("ORGMAIL") = SINPUTX



			GENCKTINI()

			GENCKTRESPO(SOPCION)

			SEGCKTPSC = GTCKTPSC()

			XOBSERV = REQUEST.FORM("OBS")

			XRTIPO  = ATABLA(1,2)
			XRCOD   = ATABLA(1,1)
			XRASG   = ATABLA(1,1)
			XRLVL   = ATABLA(1,4)
			XRSCD   = ATABLA(1,5)
			XRID   = ATABLA(1,6)

			SEGCKTPSCSIG = GTCKTPSC() & "A"
			'response.write "1 GENCKTPSC" & XRID & "<br>"
			WXX = GENCKTPSC(SEGCKTPSC,SEGCKTPSCSIG,XMODEL,"SI",XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XRLVL,XRSCD,XRID)

			SSS = GENACCION(RS(5),"INI")

			TSHW           = TSHW + 1
			ATABSW(TSHW,1) = XPDSC
			ATABSW(TSHW,2) = 1
			ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)



			''''' ANALISIS PASO SIGUIENTE


			StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
			Set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs2.eof Then
				XMODEL  = RS2(5)







				If  ATABLA(0,1) =  0 Or session("WKFCLASE") = "E" Then

					''''''''''' GENERA EL PASO FINAL AL DESTINATARIO OK

					XRTIPO  = ATABLA(2,2)
					XRCOD   = ATABLA(2,1)
					XRLVL   = ATABLA(2,4)
					XRSCD   = ATABLA(2,5)
					XRID   = ATABLA(2,6)

					XPASCOD = RS2(12)
					XPDSC   = RS2(6) & ": " & RS2(7)
					XRASG   = "NO"

					If XRTIPO = "USR" Or XRTIPO = "RSP" Then
						XRASG   = XRCOD
						If XRTIPO = "RSP" Then XRTIPO = "USR"
					End If


					If  session("WKFCLASE") = "E" Then


						XRTIPO  = RS2(15)
						XRCOD   = RS2(16)
						XRASG   = "NO"

						If XRTIPO = "USR" Or XRTIPO = "RSP" Then
							XRASG   = XRCOD
							If XRTIPO = "RSP" Then
								XRTIPO = "USR"
								For YYY=1 To ATABLA(0,0)

									If ATABLA(YYY,0) =  XRCOD Then
										XRCOD   = ATABLA(YYY,1)
										XRASG   = XRCOD
										XRLVL   = ATABLA(YYY,4)
										XRSCD   = ATABLA(YYY,5)
										XRID   = ATABLA(YYY,6)
										Exit For
									End If

								Next

							End If
						End If




					End If







					'response.write "2 GENCKTPSC" & XRID & "<br>"
					WXX = GENCKTPSC(SEGCKTPSCSIG,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD,XRID)

					TSHW           = TSHW + 1
					ATABSW(TSHW,1) = XPDSC
					ATABSW(TSHW,2) = 0
					ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

					' '''''''''''''''''''
					' Envio de mail

					XXXX = GENCKTMAIL(XRTIPO,XRCOD,mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),XRLVL)

					'''''''''''''''''''

				Else


					XSENNIV = 0
					XSENLIS = 0


					If ATABLA(3,1) =  SESSION("USUARIO") Then

						''''''''''''''' GENERAR EL PRIMER PASO OK - NO REQUIERE AUTORIZANTES PARA EL ORIGINANTE

						XSENNIV = 0

						XRTIPO  = ATABLA(3,2)
						XRCOD   = ATABLA(3,1)
						XPASCOD = RS2(12)
						XPDSC   = RS2(6) & ": " & RS2(7)
						XRASG   = ATABLA(3,1)
						XRLVL   = ATABLA(3,4)
						XRSCD   = ATABLA(3,5)
						XRID   = ATABLA(3,6)

						SEGCKTPSCSIG2 = GTCKTPSC() & "B"
						'response.write "3 GENCKTPSC" & XRID & "<br>"
						WXX = GENCKTPSC(SEGCKTPSCSIG,SEGCKTPSCSIG2,XMODEL,"SI",XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XRLVL,XRSCD,XRID)

						TSHW           = TSHW + 1
						ATABSW(TSHW,1) = XPDSC
						ATABSW(TSHW,2) = 1
						ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

						If  ATABLA(0,1) > 1 Then

							''''''''''''''''''' SI HAY LISTA DE AUTORIZANTES PARA SECTORES AFECTADOS ACTIVA PASO RECURRENTE PARA EL PRIMERO

							XSENLIS = 0

							XRTIPO  = ATABLA(4,2)
							XRCOD   = ATABLA(4,1)
							XPASCOD = RS2(12)
							XPDSC   = RS2(6) & ": " & RS2(7)
							XRASG   = "NO"
							If XRTIPO = "USR" Then  XRASG   = XRCOD
							XRLVL   = ATABLA(4,4)
							XRSCD   = ATABLA(4,5)
							XRID   = ATABLA(4,6)


						Else

							''''''''''''''''''' SI NO HAY LISTA DE AUTORIZANTES PARA SECTORES AFECTADOS ACTIVA PASO FINAL AL DESTINATARIO

							XSENLIS = 1

							SSS = GENACCION(XMODEL,"APR")

							StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
							Set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If Not rs3.eof Then
								XMODEL  = RS3(5)

								XRTIPO  = ATABLA(2,2)
								XRCOD   = ATABLA(2,1)
								XPASCOD = RS3(12)
								XPDSC   = RS3(6) & ": " & RS3(7)
								XRASG   = "NO"
								If XRTIPO = "USR" Then  XRASG   = XRCOD
								XRLVL   = ATABLA(2,4)
								XRSCD   = ATABLA(2,5)
								XRID   = ATABLA(2,6)
							End If
							Call ofv.cerrarconsulta(rs3)


						End If

						'''''''''''''' GENERAR EL SIGUIENTE PASO OK
						'response.write "4 GENCKTPSC" & XRID & "<br>"
						WXX = GENCKTPSC(SEGCKTPSCSIG2,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD,XRID)

						TSHW           = TSHW + 1
						ATABSW(TSHW,1) = XPDSC
						ATABSW(TSHW,2) = 0
						ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

						'''''''''''''''''''
						' Envio de mail
						XXXX = GENCKTMAIL(XRTIPO,XRCOD,mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),XRLVL)

						'''''''''''''''''''



					Else

						''''''''''''''' GENERAR EL PRIMER PASO AUTORIZACION PARA EL ORIGINANTE OK

						XSENNIV = 1

						If  ATABLA(0,1) > 1 Then

							XSENLIS = 0

						Else

							XSENLIS = 1

						End If

						XRTIPO  = ATABLA(3,2)
						XRCOD   = ATABLA(3,1)
						XPASCOD = RS2(12)
						XPDSC   = RS2(6) & ": " & RS2(7)
						XRASG   = "NO"
						If XRTIPO = "USR" Then  XRASG   = XRCOD
						XRLVL   = ATABLA(3,4)
						XRSCD   = ATABLA(3,5)
						XRID   = ATABLA(3,6)
						'response.write "5 GENCKTPSC" & XRID & "<br>"
						WXX = GENCKTPSC(SEGCKTPSCSIG,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD,XRID)

						TSHW           = TSHW + 1
						ATABSW(TSHW,1) = XPDSC
						ATABSW(TSHW,2) = 0
						ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

						'''''''''''''''''''
						' Envio de mail
						XXXX = GENCKTMAIL(XRTIPO,XRCOD,mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),XRLVL)

						'''''''''''''''''''


					End If

					''''''''  ACA AGREGAR TRATAMIENTO PARA ENVIO CON ORDEN INDISTINTO ''''''''''''''''''''
					'response.write SMODAL & "            " &  XSENLIS

					If SMODAL = "N" And XSENLIS = 0 Then

						XSENSTRT = 2

						If  XSENNIV = 1 Then XSENSTRT = 1

						XIND = XSENSTRT + 2

						If  ATABLA(0,1) > XSENSTRT Then

							For XIND= (XIND+1) To 200

								If ATABLA(XIND,0) <>  "L" Then Exit For

								''        XIND = XIND + 1


								''''''''''''''''' SI HAY LISTA DE AUTORIZANTES PARA SECTORES AFECTADOS ACTIVA PASO RECURRENTE PARA CADA UNO

								XRTIPO  = ATABLA(XIND,2)
								XRCOD   = ATABLA(XIND,1)
								XPASCOD = RS2(12)
								XPDSC   = RS2(6) & ": " & RS2(7)
								XRASG   = "NO"
								If XRTIPO = "USR" Then  XRASG   = XRCOD
								XRLVL   = ATABLA(XIND,4)
								XRSCD   = ATABLA(XIND,5)
								XRID   = ATABLA(XIND,6)

								SEGCKTPSCSIGX = GTCKTPSC() & "IND" & XIND

								''' GENERAR EL SIGUIENTE PASO OK
								' response.write "6 GENCKTPSC" & XRID & "<br>"
								WXX = GENCKTPSC(SEGCKTPSCSIGX,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD,XRID)

								TSHW           = TSHW + 1
								ATABSW(TSHW,1) = XPDSC
								ATABSW(TSHW,2) = 0
								ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

								'''''''''''''''''''
								' envio de mail
								XXXX = GENCKTMAIL(XRTIPO,XRCOD,mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),XRLVL)

								'''''''''''''''''''
							Next


						End If

					End If

					''''''''  FIN AGREGAR TRATAMIENTO PARA ENVIO CON ORDEN INDISTINTO ''''''''''''''''''''


				End If

			End If
			Call ofv.cerrarconsulta(rs2)

			''''' FIN ANALISIS PASO SIGUIENTE





		End If
		''' FIN VER SI CONFIRMA ENVIO


		Call ofv.cerrarconsulta(rs)

		sImagen0 = "<FONT FACE=corflaticon SIZE=2>F</FONT>"
		sImagen1 = "<FONT FACE=corflaticon SIZE=2>ñ</FONT>"
		sImagen1N = "<FONT FACE=corflaticon COLOR=#FFFFFF SIZE=2>l</FONT>"
		sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
		sImagen3 = "<FONT FACE=corflaticon SIZE=2>2</FONT>"
		sImagen4so = "<B><FONT FACE=CorFlaticon COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
		sImagen4si = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0128;</FONT></B>"
		sImagen4sA = "<B><FONT FACE=CorFlaticon COLOR=#BBBBBB SIZE=2>&#0128;</FONT></B>"
		sImagen4n = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=2>x</FONT></B>"
		sImagen6S = "<B><FONT FACE=corflaticon COLOR=#008000 SIZE=2>&#0174;</FONT></B>"
		sImagen6SS = "<B><FONT FACE=CorFlaticon COLOR=#f7aa22 SIZE=3>G</FONT></B>"
		sImagen6N = "<B><FONT FACE=corflaticon COLOR=#AA0000 SIZE=3>3</FONT></B>"
		sImagen6OT = "<B><FONT FACE=CorFlaticon COLOR=#008000 SIZE=3>&#0254;</FONT></B>"
		sImagen6T = "<B><FONT FACE=CorFlaticon COLOR=#336699 SIZE=3>q</FONT></B>"
		sImagenut = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>&#0192;</FONT></B>"
		sImagendu = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>6</FONT></B>"
		sImagenfi = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>R</FONT></B>"
		sImagenC = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
		sImagenS = "<B><FONT FACE=corflaticon COLOR=#008000 SIZE=3>R</FONT></B>"
		sImagenES = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>&#0232;</FONT></B>"
		sImagenP = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>2</FONT></B>"
		sImagenE = "<B><FONT FACE=corflaticon COLOR=#000080 SIZE=3>&#0237;</FONT></B>"
		sImagenPC = "<B><FONT FACE=CorFlaticon COLOR=#6699cc SIZE=3>.</FONT></B>"

		sInput = "SIN ASIGNAR"

		If session("NCliente") <> "NO" Then

			If session("TCliente") = "RSK" Then

				strsql = "SELECT * FROM  ORRSKCIRCUITO "
				strsql = strsql & " Where OBJID = 0" & session("NCliente")
				Set xrs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
				If Not xrs.eof Then
					If XRS("APLIC") = "OR" Then
						TIPOOBJ = "CIRCUITO AUTOEVALUACION"
					Else
						TIPOOBJ = "PLAN DE CONTINGENCIAS"
					End If

					Strsql = "Select * From objetos where codobjeto = 0" & XRS("CODOBJETOBS")
					Set xrs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
					If Not xrs.eof Then
						TIPOOBJ2 = XRS(2)
						sInput = "<B>" & XRS(3) & "</B>"
					End If


				End If
				Call ofv.cerrarconsulta(xrs)
			ElseIf session("TCliente") = "PDCORXY" Then

				If sacc="PEDIDOOK" Then

					SQL= "SELECT   'Solicitud nro:' + cast (a.objid as nvarchar)  + ': ' + BA.NOMBRE FROM PEDIDOS A "
					SQL= SQL  & "  INNER JOIN PEDIDOSDET BA ON(BA.IDPEDIDO=A.OBJID)"
					SQL= SQL  & "  INNER JOIN TIPOBJETO B ON(B.TIPOBJETO=A.TIPOBJETO)"
					SQL= SQL  & " WHERE A.OBJID=0" & session("NCliente")
					'RESPONSE.WRITE SQL
					'TIPOBJETO WHERE TIPOBJETO='" & TIPOBJETO & "'"
					Set rsx1 = ofv.crearconsultaEx(SQL,cn3,1,parametros)
					If Not rsx1.eof Then
						SESSION("NOMREQ")=UCase(rsx1(0) & "")
					End If
					Call ofv.cerrarconsulta(rsx1)

				End If

			Else

				StrSql = "Select * From OBJETOS where CODOBJETO = " & session("NCliente")
				Set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
				If Not rs.eof Then
					sInput = "<B>" & RS(3) & "</B>"
					TIPOOBJ2 = RS(2)

					Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
					Set rsX1 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
					If Not rsX1.eof Then
						TIPOOBJ = RSX1(1)
					End If
					Call ofv.cerrarconsulta(rsx1)


					SESSION("NOMREQ") = TIPOOBJ & " " & RS(2)
				End If

			End If

			SESSION("NOMREQ") = TIPOOBJ & " " & TIPOOBJ2

		End If

		sImagen = "<FONT FACE=corflaticon SIZE=4>4</FONT>"


		filas  = ""



		Set cn2 = ofv.conectar(ofv.strconn0)

		StrSql = "Select * From USUARIOS WHERE ID = '" & Session("Usuario") & "' "
		Set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
		If Not rs.eof Then

			sInput = "CIRCUITO " & session("WKFDSC")
			celdas = ""
			''''   sInput = session("WKFDSC")
			celdas = celdas & ofv.GenCelda("","ttr colspan=2 ","","80%","","","",sInput,"si")
			filas  = filas & ofv.GenRow("","","","","22","","",celdas,"si")


			sInput = "ORIGINANTE:"
			celdas = ofv.GenCelda("","tt","","","","","",sInput,"si")

			sInput = sImagen4so & "&nbsp;" & Session("SEGUsuario") & ""
			celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")

			filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

			sInput = "GRUPO:"
			celdas = ofv.GenCelda("","tt","","","","","",sInput,"si")

			sgrp = ""

			StrSql = "Select * From grupos WHERE ID = '" & Session("SEGGRP") & "' "
			Set rs5 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			If Not rs5.eof Then
				sgrp = rs5(1)
			End If

			sInput = sImagen6S & "&nbsp;" & SESSION("SEGGRP") & " - " & sgrp & ""
			celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")

			filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

			sInput = "SECTOR:"
			celdas = ofv.GenCelda("","tt","","","","","",sInput,"si")

			sInput = "SIN ASIGNAR"

			Set cn2 = ofv.conectar(ofv.strconn2)
			On Error Resume Next
			strsql = "Select * From ECO where ECO = " & rs(21)
			Set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			If Not rsniv.eof Then sInput = rsniv(1)
			Call ofv.cerrarconsulta(rsniv)
			Call ofv.cerrarconn(cn2)

			celdas = celdas & ofv.GenCelda("","uti","","","","","",sImagen6SS & "&nbsp;" & sInput,"si")
			filas = filas & ofv.GenRow("","","","","","","",celdas,"si")




		End If



		celdas = ofv.GenCelda("","vt colspan=2 ","","","","","","&nbsp;","si")
		filas = filas & ofv.GenRow("","","","","","","",celdas,"si")
		TAB = TAB  & ofv.GenTabla("","","","90%","","0","","0","0","","",filas,"si") & "<BR>"



		If act = "S" Then

			filas = ""
			For ISHW = 1 To TSHW

				sinput = sImagenPC  & "&nbsp;PASO  " & ATABSW(ISHW,1)
				celdas = ofv.GenCelda("","vT ","","","50%","","",sInput,"si")

				If ATABSW(ISHW,2) = 0 Then
					SINPUT3 = "<FONT SIZE=1 COLOR=#008000><U>INICIADO</U></FONT>"
				Else
					SINPUT3 = "<FONT SIZE=1 COLOR=#993333><U>CONFORMADO</U></FONT>"
				End If

				sinput = sImagenC & "&nbsp;ESTADO:&nbsp;" & SINPUT3 & "<BR>" & GETDH()
				celdas = celdas & ofv.GenCelda("","vt","","","50%","","",sInput,"si")

				filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")


				celdas = ofv.GenCelda("","vt COLSPAN=2","","","50%","","","&nbsp;","si")
				filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")

				sinput = "DESTINO:"
				celdas = ofv.GenCelda("","vt","","","50%","","",sInput,"si")

				sinput =  sImagen6SS & "&nbsp;" & ATABSW(ISHW,3)
				celdas = celdas & ofv.GenCelda("","vt","","","50%","","",sInput,"si")

				filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")

				celdas = ofv.GenCelda("","vt COLSPAN=2 style='border-TOP-style:SOLID;border-TOP-width:1;border-TOP-color:#6699CC;' ","","","50%","","","&nbsp;","si")
				filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")


			Next

			sinput = "&nbsp;"
			celdas = ofv.GenCelda("","vt","","","50%","","",sInput,"si")

			sinput = "<a class=tt  href=EDREQUER.asp?OPCION=RELOAD"
			sinput = sinput  & " ><u><b>VOLVER A SOLICITUD</b></u></a>"
			celdas = CELDAS & ofv.GenCelda("","tt","CENTER","","50%","","",sInput,"si")

			''' filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

			TAB = TAB  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")
			If SACC="PEDIDOOK" Then
				sInput="<B>CIRCUITO WORKFLOW DE SOLICITUD ABM INICIADO Y ENVIADO AL IMPLEMENTADOR</B><br>" '&  sacc & "  "  & session("NCliente") & " " & session("TCliente")
				' session("NCliente") <> "NO" THEN

				'IF session("TCliente") = "RSK" THEN

				'
				celda =  ofv.GenCelda("","tt style='text-align: center !important;border-style:solid;border-width:1;border-color:#dddddd;'","","","50%","","",sInput,"si")
				fila=ofv.GenRow("","","","","","","",celda,"si")
				TAB =  ofv.GenTabla("","","","50%","","0","","0","0","","",fila,"si")
			End If

		Else

			sinput = sImagenPC  & "&nbsp;PASO  " & XPDSC1
			celdas = ofv.GenCelda("","vT","","","50%","","",sInput,"si")

			sinput = sImagenC & "&nbsp;ESTADO:&nbsp;PENDIENTE DE INICIO"
			celdas = celdas & ofv.GenCelda("","vt","","","50%","","",sInput,"si")

			filas =  ofv.GenRow("","","","","","","",celdas,"si")

			TAB = TAB  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


			Celda = ofv.GenCelda("","vt colspan=3","CENTER","100%","","","","OBSERVACIONES","si")
			FilaS = Ofv.GenRow("","","","","","","",Celda,"si")

			nCol = 80

			wdescri = "EL AUTORIZANTE NO REGISTRA OBSERVACIONES "

			t3b = "<font class=at><textarea class=dt "
			t3b = t3b & "name=OBS cols=" & nCol & " rows=10 "
			If bModificar Then
				'''     t3b = t3b & " onblur=Usuarios.submit() "
				t3b = t3b & " onkeypress=" & Chr(34)
				t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
				t3b = t3b & " extenso'); return false;}" & Chr(34) & " >"
			Else
				t3b = t3b & " readonly >"
			End If
			xdatos =  wdescri
			t3b = t3b & xdatos & "</textarea></font><br><br>"
			Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","",t3b,"si")
			FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

			sinput = "&nbsp;"
			celda = ofv.GenCelda("","","center","","50%","","",sInput,"si")

			'' Agregado 20042015
			'''sinput = "<input CLASS=BT type=submit value='CONTINUAR' size=10 >"
			sinput = "<input CLASS=BT type=submit value='CONTINUAR' size=10 onclick=""this.style.cursor='wait';this.style.visibility='hidden';""  >"
			'' Agregado 20042015 FIN

			'celda = celda & ofv.GenCelda("","","center","","1%","","",sInput,"si")
			XVUELTA=REQUEST.QUERYSTRING
			'response.write xvuelta


			If wkaplic="OR" Then

				SFUNC="document.Usuarios.action='EdObjwFlor.ASP?" & XVUELTA &"';document.Usuarios.submit();"
			Else
				' sAccion  = "edobjwfl.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo
				SFUNC="document.Usuarios.action='edobjwfl.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo  &"';document.Usuarios.submit();"
			End If


			sInput2 = ofv.GenerarInput("","VOLVER","button","","10","bt",sFunc)

			If bModificar Then
				sAccion = "EDWKFSTART.ASP?CantRegMover=0&OPCION=" & SOPCION & "&ACT=S&MODAL=" & SMODAL
				sinput = ofv.GenForm("Usuarios","",sAccion,"","post","",t3b & sinput & sinput2,"si")
			Else
				sinput = ofv.GenForm("","","","","post","",t3b & sinput & sinput2,"si")
			End If


			celda = celda & ofv.GenCelda("","abc colspan=3","center","","1%","","",sInput ,"si")

			''      sinput = "&nbsp;"
			''     celda = celda & ofv.GenCelda("","","center","","50%","","",sInput,"si")
			formulario0 =  ofv.GenRow("","","","","","","",Celda,"si")




			TAB = TAB  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario0,"si") & "<BR>"


		End If



		Celda = ofv.GenCelda("","","center","","","","",TAB,"si")
		FilaS = ofv.GenRow("","","","","","","",Celda,"si")

		sHTML = sHTML & "<DIV CLASS='DINDC' >"
		sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",filas,"si")
		sHTML = sHTML & "</DIV>"



	Else
		SHTML = SHTML & "<P>ERROR EN EL INICIO DEL CIRCUITO</P>"

	End If





	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)



	If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

		sAccion = Session("ASPSTR") & "&ACCIONEJEC=ABREESTAD"
		Retornar = ofv.BotonVolver(sAccion)

		xxxxx = ofv.botoneraESSX("0","","","","","","","","","no","",Retornar,"")


	End If

	'--------------------------------------------------------------------------------
	sHTML = sHTML & "</div>"

	Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")

	'--------------------------------------------------------------------------------

	sHTML = sHTML & xxxxx


	sHTML = sHTML  & "</body></html>"
End If

response.write sHTML

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function GENCKTINI()

	SEGCKT = GTCKT()

	SESSION("SEGCKT") = SEGCKT

	strsql = "insert into SEGCIRCUITO (MODEL,COMIENZO,FIN,OBJTIPO,OBJCOD,SEGCKT,RESPONSABLE,CKTDSC,CKTMODAL) values "
	STRSQL = STRSQL & " ('" & SESSION("WKF") & "','" & GETDH() & "','NO','" & session("TCliente") & "',"
	STRSQL = STRSQL & "'" & session("NCliente") & "','" & SEGCKT & "','" & SESSION("USUARIO")
	STRSQL = STRSQL & "','" & session("WKFDSC") & "','" & sModal & "')"
	Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

End Function


Function GENCKTRESPO(XTIPO)

	''' session("SegNivel")

	ATABLA(0,0) =  2
	ATABLA(0,1) =  0

	''''''''' 1. GENERA ORIGINANTE DEL CIRCUITO --> USUARIO DE LA SESION.

	ATABLA(1,0) =  "O"
	ATABLA(1,1) =  SESSION("USUARIO")
	ATABLA(1,2) =  "USR"
	ATABLA(1,3) =  "NO"
	ATABLA(1,4) =  session("SegNivel")
	ATABLA(1,5) =  "0"
	ATABLA(1,6) =  0
	xxx = getinfouser(SESSION("USUARIO"))
	'response.write "1" & "<br>"
	STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,SEGEMAIL) values "
	STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','NO','O','USR','" & SESSION("USUARIO") & "','10','" & session("SegNivel") & "','0','" & session("SGU_MAIL") & "')"
	'response.write strsql & "<br>"
	Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

	If session("WKFCLASE") <> "E" Then

		''''''''' 2. GENERA DESTINATARIO DEL CIRCUITO --> A: AUTO > SECTOR INTERVINIENTE - NIVEL DESTINO. B: APRO > USUARIO SOLICITANTE - NIVEL 9

		'  IF XTIPO = "AUTO" THEN

		ATABLA(2,0) =  "D"
		ATABLA(2,1) =  SESSION("USUARIO")
		ATABLA(2,2) =  "USR"
		ATABLA(2,3) =  "NO"
		ATABLA(2,4) =  session("SegNivel")
		ATABLA(2,5) =  "0"
		ATABLA(2,6) =  999
		xxx = getinfouser(SESSION("USUARIO"))
		STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,sector,segmail) values "
		STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','NO','D','USR','" & SESSION("USUARIO") & "','999','"
		STRSQL = STRSQL & session("SegNivel") & "','0','" & session("CLIEMP") & "','" & session("UsrSector") & "','"  &  session("SGU_MAIL")   & "')"
		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)


		'  END IF




		STRSQL = "Select * From WKFCKTMASTER WHERE TIPO = 'R' AND WKFCOD = '" & SESSION("WKF") & "' ORDER BY ORDEN "
		Set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		Do Until rs3.eof

			XORDEN = 10

			XINDT = 2

			ATABLA(0,0) =  2
			ATABLA(0,1) =  0

			' DESLEVL = 9
			' SUPLEVL = 9


			''''''''' 3. GENERA AUTORIZANTE DEL ORIGINANTE DEL CIRCUITO --> A: AUTO > SECTOR SOLICITANTE - NIVEL MINIMO SUPERVISION.
			'                                                               B: APRO > SECTOR INTERVINIENTE - NIVEL MINIMO SUPERVISION.
			'                                                                 (SOLO SI HAY PASOS RECURRENTES DE AUTORIZACION)

			XINDT = XINDT + 1

			'   IF XTIPO = "AUTO" THEN
			If SUPLEVL < session("SegNivel") Then

				XXXX = GENSECTORUSR(SESSION("USUARIO"))

				XORDEN = XORDEN + 10


				ATABLA(XINDT,0) =  "L"
				ATABLA(XINDT,1) =  SESSION("SECTORIG")
				ATABLA(XINDT,2) =  "SEC"
				ATABLA(XINDT,3) =  RS3(5)
				ATABLA(XINDT,4) =  SUPLEVL
				ATABLA(XINDT,5) =  SESSION("SECTGEREN")
				ATABLA(XINDT,6) =  XORDEN
				strsql = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,segmail) values "
				STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','SEC','" & SESSION("SECTORIG") & "','" & XORDEN & "','" & SUPLEVL & "','" & SESSION("SECTGEREN") & "','','NO')"
				Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Else

				XORDEN = XORDEN + 10

				xxx = getinfouser(SESSION("USUARIO"))
				STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,sector,SEGEMAIL) values "
				STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','USR','" & SESSION("USUARIO") & "','" & XORDEN & "','" & session("SegNivel") & "','0','" & session("CLIEMP") & "','" & session("UsrSector") & "','" & session("SGU_MAIL")  & "')"
				Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

				ATABLA(XINDT,0) =  "L"
				ATABLA(XINDT,1) =  SESSION("USUARIO")
				ATABLA(XINDT,2) =  "USR"
				ATABLA(XINDT,3) =  RS3(5)
				ATABLA(XINDT,4) =  session("SegNivel")
				ATABLA(XINDT,5) =  "0"
				ATABLA(XINDT,6) =  xorden
			End If



			'  END IF


			''''''''' 4. GENERA LISTA DE AUTORIZANTES PARA SECTORES AFECTADOS DEL CIRCUITO --> AUTO Y APRO >
			'                                                                 SECTOR AFECTADO - NIVEL MINIMO SUPERVISION.
			'                                                                  (SOLO SI HAY PASOS RECURRENTES DE AUTORIZACION)


			'STRSQL = "Select * From OBJWFLDET where CODOBJETO = " & session("NCliente") & "  And "
			'STRSQL = STRSQL & " TIPOBJETO = '" & session("TCliente") & "'  And "
			'STRSQL = STRSQL & " CKTTYPE = '" & session("CKTTYPE") & "'  ORDER BY ORDEN,OBJID "


			STRSQL = "Select A.* From OBJWFLDET A  "
			STRSQL = STRSQL & " INNER JOIN Objwfl B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND B.ESQAUTO =1) "
			STRSQL=  STRSQL & " where A.CODOBJETO = " & session("NCliente") & "  And "
			STRSQL = STRSQL & " A.TIPOBJETO = '" & session("TCliente") & "'  And "
			STRSQL = STRSQL & " A.CKTTYPE = '" & session("CKTTYPE") & "'  ORDER BY A.ORDEN,A.OBJID "



			''     response.write strsql
			Set rs1 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			Do Until rs1.eof
				XORDEN = XORDEN + 10
				XINDT = XINDT + 1

				XXXX = GENSECTORUSR(RS1(4))






				ATABLA(XINDT,0) =  "L"
				ATABLA(XINDT,1) =  RS1(4)
				ATABLA(XINDT,2) =  "USR"
				ATABLA(XINDT,3) =  RS3(5)
				ATABLA(XINDT,4) =  SUPLEVL
				ATABLA(XINDT,5) =  SESSION("SECTORIG")
				ATABLA(XINDT,6) =  xorden
				XX = getinfouser(rs1(4))

				STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,sector,SEGEMAIL) values "
				STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','USR','" & RS1(4) & "','" & XORDEN & "','" & SUPLEVL & "','" & SESSION("SECTORIG") & "','" & rs1(14) & "','','" & session("SGU_MAIL")  & "')"

				Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

				RS1.MOVENEXT

			Loop



			If session("TCliente") = "RSK" Then
			ElseIf sacc="PEDIDOOK" Then



				sql="select GRPUSR from pedidoSDKdet  where codobjeto=0 and tipobjeto='" & session("TCliente") & "' order by objid "


				' response.write sql  & "<br>" & chr(13)
				Set rx =  ofv.crearconsultaEx(sql,CN3,1,volver)

				Do While Not  rx.EOF
					XORDEN = XORDEN + 10
					XINDT = XINDT + 1
					' XUSUC=mid( rx(4)& "" ,5,len(rx(4)) &"" )
					P_WKFAPRO=Mid( rx(0)& "" ,5,Len(rx(0)) &"" )
					XXXX = GENSECTORUSR(P_WKFAPRO)


					ATABLA(XINDT,0) =  "L"
					ATABLA(XINDT,1) =  P_WKFAPRO
					ATABLA(XINDT,2) =  "USR"
					ATABLA(XINDT,3) =  RS3(5)
					ATABLA(XINDT,4) =  SUPLEVL
					ATABLA(XINDT,5) =  SESSION("SECTORIG")
					ATABLA(XINDT,6) =  xorden
					XX = getinfouser(P_WKFAPRO)
					'response.write "6" & "<br>"
					STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,SEGEMAIL) values "
					STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','USR','" &  P_WKFAPRO & "','" & XORDEN & "','" & SUPLEVL & "','" & SESSION("SECTORIG") & "','" & session("SGU_MAIL") & "')"
					' RESPONSE.WRITE STRSQL & "<br>" & chr(13)
					'	     response.write strsql & "<br>"
					Call ofv.crearconsultaEx(StrSql,cn,1,parametros)
					RX.MOVENEXT

				Loop
				Call ofv.cerrarconsulta(rx)

			Else

				strsql = "select * from Defwfl "
				Set rx =  ofv.crearconsultaEx(StrSql,CN3,1,volver)
				If Not rx.eof Then
					xactiva = rx(1)
				End If
				Call ofv.cerrarconsulta(rx)

				If xactiva Then
					'  strsql = "select * from objwfldet Where codobjeto = 0"


					'   IF session("PASOOBS")="OBS" THEN
					'   strsql = strsql & " and tipobjeto ='WKF00016631121722600921928999'  order by orden, objid"
					'  else
					'   strsql = strsql & " and tipobjeto ='WKF00016631121722600921928'  order by orden, objid"
					'  end if

					strsql =  " select * from objwfldet  "
					strsql = strsql & "  INNER  join tipObjeto on (tipobjeto.tipobjeto='" & session("TCliente") & "' )"
					If session("EMPMUN") <> "S" Then
						strsql = strsql & "     Where (objwfldet.codobjeto = 0 OR  objwfldet.TIPOBJETO='ALL') and codemp=" & session("UN")
					Else
						strsql = strsql & "     Where (objwfldet.codobjeto = 0 OR  objwfldet.TIPOBJETO='ALL') and codemp=" & session("UN")
					End If
					If session("PASOOBS")="OBS" Then
						strsql = strsql & " AND NOT  tipObjeto.WKFAUTO=objwfldet.TIPOBJETO"
					Else
						strsql = strsql & " AND NOT  tipObjeto.WKFAPRO=objwfldet.TIPOBJETO"
					End If
					strsql = strsql & "  order by orden, objid "

					''                response.write strsql
					Set rx =  ofv.crearconsultaEx(StrSql,CN3,1,volver)
					If Not rx.eof Then
						Do Until rx.eof


							XORDEN = XORDEN + 10
							XINDT = XINDT + 1

							XXXX = GENSECTORUSR(rx(4))


							ATABLA(XINDT,0) =  "L"
							ATABLA(XINDT,1) =  rx(4)
							ATABLA(XINDT,2) =  "USR"
							ATABLA(XINDT,3) =  RS3(5)
							ATABLA(XINDT,4) =  SUPLEVL
							ATABLA(XINDT,5) =  SESSION("SECTORIG")
							ATABLA(XINDT,6) =  xorden
							xx = getinfouser(rx(4))
							STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,sector,SEGMAIL) values "
							STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','USR','" & rx(4) & "','" & XORDEN & "','" & SUPLEVL & "','" & SESSION("SECTORIG") & "','" & rx(14) & "','','" & session("SGU_MAIL") & "')"
							''       response.write  STRSQL
							Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

							rx.movenext
						Loop
					End If
					Call ofv.cerrarconsulta(rx)
					' incidente nro 1374

					If session("EMPMUN") <> "S" Then
						strsql =  " select * from objsdkdet  where codobjeto=0 And tiposeg ='0' and tipobjeto='" & session("TCliente") &  "' and codemp=" & session("UN")
					Else
						strsql =  " select * from objsdkdet  where codobjeto=0 And tiposeg ='0' and tipobjeto='" & session("TCliente") & "' and codemp=" & session("UN")
					End If
					strsql = strsql & "  order by  objid "

					' response.write strsql
					Set rx =  ofv.crearconsultaEx(StrSql,CN3,1,volver)
					If Not rx.eof Then
						Do Until rx.eof


							XORDEN = XORDEN + 10
							XINDT = XINDT + 1
							XUSUC=Mid( rx(4)&"" ,5,Len(rx(4)) &"" )
							XXXX = GENSECTORUSR(XUSUC)


							ATABLA(XINDT,0) =  "L"
							ATABLA(XINDT,1) = XUSUC
							ATABLA(XINDT,2) =  "USR"
							ATABLA(XINDT,3) =  RS3(5)
							ATABLA(XINDT,4) =  SUPLEVL
							ATABLA(XINDT,5) =  SESSION("SECTORIG")
							ATABLA(XINDT,6) =  xorden
							xxx = getinfouser(XUSUC)
							STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,sector,SEGMAIL) values "
							STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','USR','" & XUSUC & "','" & XORDEN & "','" & SUPLEVL & "','" & SESSION("SECTORIG") & "','" & rx(5) & "','','" & session("SEG_MAIL") & "')"
							''   response.write  STRSQL
							Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

							rx.movenext
						Loop
					End If
					Call ofv.cerrarconsulta(rx)










				End If

			End If

			RS3.MOVENEXT
		Loop

		If XINDT > 0 Then

			ATABLA(0,0) =  XINDT
			ATABLA(0,1) =  XINDT - 2

		End If

	Else

		''''''''' E2. GENERA ORIGINANTE DEL CIRCUITO COMO SOLICITANTERO.

		ATABLA(2,0) =  "SOLICITANTERO"
		ATABLA(2,1) =  SESSION("USUARIO")
		ATABLA(2,2) =  "USR"
		ATABLA(2,3) =  "NO"
		ATABLA(2,4) =  "0" & session("SegNivel")
		ATABLA(2,5) =  "0"

		STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,sector,SEGMAIL) values "
		STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','NO','SOLICITANTERO','USR','" & SESSION("USUARIO") & "','20','0" & session("SegNivel") & "','0','" & SESSION("CLIEMP") & "','" & session("UsrSector") & "','" & session("UsrMail")  & "')"
		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)


		XORDEN = 20

		XINDT = 2

		ATABLA(0,0) =  2
		ATABLA(0,1) =  1



		''''''''' E3. GENERA SET DE AUTORIZANTES PARA RESPONSABLES PREDETERMINADOS (CIRCUITO TIPO:EFECTIVO)


		STRSQL = "Select A.* From OBJWFLDET "

		STRSQL = STRSQL & " INNER JOIN Objwfl B ON (B.CODOBJETO=A.CODOBJETO AND B.TIPOBJETO=A.TIPOBJETO AND B.ESQAUTO =1) "
		STRSQL = " where A.CODOBJETO = " & session("NCliente") & "  And "
		STRSQL = STRSQL & " A.TIPOBJETO = '" & session("TCliente") & "'  And "
		STRSQL = STRSQL & " A.CKTTYPE = '" & session("CKTTYPE") & "'  ORDER BY ORDEN,OBJID "
		Set rs1 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
		Do Until rs1.eof
			XORDEN = XORDEN + 10
			XINDT = XINDT + 1

			XXXX = GENSECTORUSR(RS1(4))






			ATABLA(XINDT,0) =  RS1(5)
			ATABLA(XINDT,1) =  RS1(4)
			ATABLA(XINDT,2) =  "USR"
			ATABLA(XINDT,3) =  "NO"
			ATABLA(XINDT,4) =  "0" & SUPLEVL
			ATABLA(XINDT,5) =  SESSION("SECTORIG")
			ATABLA(XINDT,6) =  xorden
			xxx = getinfouser(RS1(4))
			STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO,codempre,sector) values "
			STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','NO','" & RS1(5) & "','USR','" & RS1(4) & "','" & XORDEN & "','0" & SUPLEVL & "','" & SESSION("SECTORIG") & "','" & rs1(14) & "','')"
			Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			RS1.MOVENEXT

		Loop

		ATABLA(0,0) =  XINDT
		ATABLA(0,1) =  XINDT - 1


	End If

	sql = "select * from segcktresp where segckt='" & SESSION("SEGCKT") & "' order by orden"
	Set rx = ofv.crearconsultaEx(sql,cn,1,parametros)
	iii = 1
	Do While Not rx.eof

		If rx("modo") = "O" Then
			ATABLA(1,6) =  rx(0)
		ElseIf rx("modo") = "D" Then
			ATABLA(2,6) =  rx(0)
		Else

			iii = (rx("orden") / 10) + 1
			ATABLA(iii,6) =  rx(0)

		End If
		'response.write "codigo: " & rx("codigo") & " modo: " & rx("modo") & " orden: " & rx("orden") & " id: " & rx(0) & " indice: " & iii & " ATABLA: " & ATABLA(iii,1) & "<br>"
		rx.movenext
	Loop

End Function

Function GENCKTMAIL(WETIPO,WECOD,WEMESSAGE,WEORIGEN,WELVL)


	XRESPO = WECOD & ""

	If XRESPO <> "" And XRESPO <> "NO" Then

		Set cn2 = ofv.conectar(ofv.strconn0)

		XPARA = GETRESPOCKT(WECOD,WETIPO)
		STRSQL = "SELECT EMAIL,DESCRIPCION FROM USUARIOS WHERE "

		XGUMENSAGEw = WEMESSAGE


		XMESAURL  = ""
		If Session("CORUNQWKF") = "S" Then

			XMESAURL = getobjurlWK(session("NCliente"),session("TCliente"),SEGCKTPSC)

		Else
			If  SESSION("URLTYP")  = "V" Then

				XMESAURL  =  Left(request.servervariables("SCRIPT_NAME"),InStr(Right(request.servervariables("SCRIPT_NAME"),Len(request.servervariables("SCRIPT_NAME"))-1),"/"))

			End If

			If SESSION("URLPORT")  = "N" Then
				XMESAURL   = "http://" & request.servervariables("SERVER_NAME")  & XMESAURL
			Else
				XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
			End If
		End If
		XGUMENSAGEw = Replace(XGUMENSAGEw,"<APLURL>","<U><A HREF='" & XMESAURL  & "' TARGET='_blank' >")
		XGUMENSAGEw = Replace(XGUMENSAGEw,"</APLURL>","</A></U>")

		WEMESSAGE = XGUMENSAGEw

		Select Case WETIPO

			Case "USR"
			' STRSQL = STRSQL & " ID = '" & XRESPO & "' "


			xxxx = getInfoUser(XRESPO)
			If session("SGU_mail")<> "" Then    SINPUTX= session("SGU_mail")


			stowho = SINPUTX & ""
			If Len(stowho) > 0 Then
				stowho = Trim(stowho)
				If sopcion = "APRO" Then
					subject = "SOLICITUD DE OBSERVACION"
				Else

					subject = "SOLICITUD DE AUTORIZACION"
				End If
				If session("TCliente") = "PDCORXY" Then
					subject = "SOLICITUD ABM PEDIDO"
				ElseIf session("TCliente") = "SN" Then
					subject = "SOLICITUD ABM NORMATIVA"
				End If

				If WETIPO = "USR" Then
					XPARA2 = XPARA
				Else
					XPARA2 = XPARA & " / " & session("SGU_nombre")
				End If

				Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
				''  message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
				'  response.write Xmessage & "       "  & stowho & "                "  &  WEORIGEN & "        " & session("MailComponent")
				Call SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
			End If

			Case "GRP"

			' STRSQL = STRSQL & " GRUPO = '" & XRESPO & "' and ( NOT nivel >  " & WELVL & " ) "

			Call  getUsrIngroup(XRESPO)
			A_SGU_UIGuserid =Split(session("SGU_UIGuserid"),Chr(9))
			A_SGU_UIGdescri =Split(session("SGU_UIGdescri") ,Chr(9))

			'   set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			' DO UNTIL RSXPA.EOF
			For a_i=0 To UBound(SGU_UIGdescri)-1
				xxxx = getInfoUser(A_SGU_UIGuserid(a_i))
				stowho = session("SGU_mail") & ""
				If Len(stowho) > 0 Then
					stowho = Trim(stowho)
					If sopcion = "APRO" Then
						subject = "SOLICITUD DE OBSERVACION"
					Else

						subject = "SOLICITUD DE AUTORIZACION"
					End If
					If WETIPO = "USR" Then
						XPARA2 = XPARA
					Else
						XPARA2 = XPARA & " / " & A_SGU_UIGdescri(a_i)
					End If
					'' message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
					Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
					Call SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
				End If

				'RSXPA.MOVENEXT
			Next

			Case "SEC"
			XRESPO = 0 & XRESPO

			STRSQL = STRSQL & " SECTOR = " & XRESPO & " and ( NOT nivel >  " & WELVL & " ) "

		End Select







	End If

End Function



Function GETRESPOCKT(RESPCOD,RESPTIP)

	GETRESPOCKT = "SIN ASIGNAR"



	XRESPO = RESPCOD & ""

	If XRESPO <> "" And XRESPO <> "NO" Then

		Set cn2 = ofv.conectar(ofv.strconn0)
		Set cn3 = ofv.conectar(ofv.strconn2)

		Select Case RESPTIP

			Case "USR"
			' strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
			' set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			' if not rsniv.eof then sInputx = rsniv(1)
			' call ofv.cerrarconsulta(rsniv)

			xxxx = getInfoUser(XRESPO)
			If session("SGU_nombre") <> "" Then    SINPUTX= session("SGU_nombre")


			Case "GRP"
			strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
			Set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			If Not rsniv.eof Then sInputx = rsniv(1)
			Call ofv.cerrarconsulta(rsniv)

			Case "SEC"
			XRESPO = 0 & XRESPO
			strsql = "Select * From ECO where ECO = " & XRESPO
			'''     response.write strsql
			Set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If Not rsniv.eof Then sInputx = rsniv(1)
			Call ofv.cerrarconsulta(rsniv)

		End Select

		GETRESPOCKT = sInputx

	End If


End Function



Function GENCKTPSC(XPSC,XSIG,XMODEL,XFIN,XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XNIVEL,XSUBCOD,XRID)
	'response.write "XPSC:" & XPSC & "<br>" & "XSIG:" & XSIG & "<br>" & "XMODEL:" & XMODEL & "<br>" & "XFIN:" & XFIN & "<br>" & "XRTIPO:" & XRTIPO & "<br>" & "XRCOD:" & XRCOD & "<br>" & "XRASG:" & XRASG & "<br>" & "XPDSC:" & XPDSC & "<br>" & "XOBSERV:" & XOBSERV & "<br>" & "XNIVEL:" & XNIVEL & "<br>" & "XSUBCOD:" & XSUBCOD & "<br>" & "<br>"
	SEGCKTPSC = XPSC

	'ESTA RUTINA CIERRA EL PASO 999 AUTOMATICO
	sqlF="select cast ( fin as int)  from WKFPASOS where wkfpascod='" & XMODEL & "'"
	Set rspsfinal=ofv.crearconsultaEx(sqlf, cn, 1, parametros)
	spsfin="N"
	If Not rspsfinal.eof Then
		If rspsfinal(0) & ""  = "1"  Then spsfin="S"
	End If
	Call ofv.cerrarconsulta(rspsfinal)

	SESSION("SEGCKTPSC") = SEGCKTPSC

	XCOM = GETDH()
	If XFIN = "SI" Then
		XFIN = GETDH()
	Else
		XFIN = "NO"
	End If



	'response.write "IDRESP: " & IDRESP & "<br>"
	salida = ""

	If XNIVEL = 3 And XSUBCOD = 0 Then
		salida = "OK"
	End If

	strsql = "insert into SEGPASO (SEGCKT,SEGCKTPSC,MODEL,COMIENZO,FIN,OKRCH,RESPTIPO,RESPONSABLE,RESPASG,PSCDSC,OBSERVACIONES,NIVEL,SUBCODIGO,SEGIDRESP,SEGSALIDA) values "
	STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & SESSION("SEGCKTPSC") & "','" & XMODEL & "','" & GETDH() & "','" & XFIN & "','" & XSIG & "','" & XRTIPO & "','" & XRCOD & "','" & XRASG & "','" & XPDSC & "','" & XOBSERV & "','" & XNIVEL & "','" & XSUBCOD & "','" & XRID & "','" & salida & "')"
	'response.write "STRSQL: " & STRSQL & "<br>"
	Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

	sql = "select iddestino from envioformdet2 where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "'"
	Set rk = ofv.crearconsultaEx(sql,cn2,1,parametros)

	'controla que tenga algun destino el formulario para entrar a la rutina de copiado de documentos
	flg = "NO"
	If Not rk.eof Then

		Do While Not rk.eof

			If rk(0) <> 0 Then
				flg = "SI"
			Else
				flg = "NO"
			End If

			rk.movenext
		Loop

	End If

	'RUTINA PARA COPIAR LOS ADJUNTOS DEL DOCUMENTO A LOS DESTINOS PREVIAMENTE SELECCIONADOS
	If session("APLMDCLI") = "BGAL" And session("COPIA") = "S" And tipocpt = "K" And spsfin = "S" And flg = "SI" Then

		xsql = "select estado from objprty where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "'"
		Set ri = ofv.crearconsultaEx(xsql,cn2,1,parametros)

		xestado = ""
		If Not ri.eof Then xestado = ri(0)

		If xestado = 3 Then

			sql = "select c.directorio,c.archivo,d.url,b.iddestino,c.tipo from objetos a "
			sql = sql & " left join envioformdet2 b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto) "
			sql = sql & " left join objhtm c on (a.codobjeto=c.codobjeto and a.tipobjeto=c.tipobjeto) "
			sql = sql & " left join destinosurl d on (b.iddestino=d.id) "
			sql = sql & " where a.codobjeto=" & session("NCliente") & " and a.tipobjeto='" & session("TCliente") & "'"
			Set ro = ofv.crearconsultaEx(sql,cn2,1,parametros)

			If Not ro.eof Then

				Dim fs
				Set fs=Server.CreateObject("Scripting.FileSystemObject")

				Do While Not ro.eof
					'controla para copiar que no sea 0 el destino cargado para que no pinche
					If CStr(ro(3)) <> "0" Then

						If ro(4) = 1 Then

							xOrigen = session("DirUpload") & ro(0) & "\" & ro(1)

						ElseIf ro(4) = 3 Then

							xOrigen = ro(1)

						End If

						xDestino = ro(2)

						fs.CopyFile xOrigen,xDestino

					End If

					ro.movenext
				Loop

				Set fs=Nothing

				asunto = "Claridad Galicia - Formulario vigente"

				sql = "select a.url,c.descri,a.mail1,a.mail2 from destinosurl a "
				sql = sql & " inner join envioformdet2 b on (a.id=b.iddestino) "
				sql = sql & " inner join objetos c on (b.codobjeto=c.codobjeto and b.tipobjeto=c.tipobjeto) "
				sql = sql & " where b.codobjeto=" & session("NCliente") & " and b.tipobjeto='" & session("TCliente") & "'"
				Set rj = ofv.crearconsultaEx(sql,cn2,1,parametros)

				cuerpo = "Se informan las urls a las que se copio el archivo adjunto del formulario: " & rj(1) & "<br><br>"
				mails = ""

				Do While Not rj.eof

					cuerpo = cuerpo & rj(0) & "<br>"
					mails = mails & sep & rj(2) & ";" & rj(3)

					sep = ";"

					rj.movenext
				Loop

				xx = SendMailAtt(cuerpo,mails,"mabraham@corconsulting.com.ar",asunto,"CDOSYS","SI","1","NO","","xClaridadGalicia@bancogalicia.com.ar",0)

			End If

		End If

	End If

End Function

Function GENACCION(XPASO,XTPP)

	'''  response.write "entro accion: " & XPASO & " -- " & XTPP & " // "
	SQL = "SELECT TIPOCPT FROM TIPOBJETO WHERE  TIPOBJETO='" & session("TCliente") & "'"
	'    response.write sql & "<br>"
	Set RS1 = ofv.crearconsultaEx(SQL, cn3, 1, parametros)
	PEDIDO="N"
	If Not RS1.EOF Then  PEDIDO=RS1(0)

	Call OFV.CERRARCONSULTA(RS1)


	StrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & XPASO & "' AND TIPO = '" & XTPP & "' "
	' response.write strsql
	Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not RS1.EOF Then
		If RS1(4) = "A" Then
			' response.write "estado: " & RS1(5) & "<br>" & session("TCliente") & "        " & session("NCliente")
			If  session("TCliente") = "RSK" Then
				GETWKFTAB = "ORRSKCIRCUITO"
			ElseIf PEDIDO="PD" Or PEDIDO="SN" Then
				GETWKFTAB = "PEDIDOS"
			Else
				GETWKFTAB = "OBJPRTY"
			End If

			XXXX = GETWKFMANAGER(session("NCliente"),session("TCliente"),RS1(5),GETWKFTAB,"WKF")
			' ffff
		End If


	End If




End Function


Function GENSECTORUSR(XUSR)

	XSECTORIG  = 0
	XSECTORIGB = XSECTORIG
	XSECTGEREN = 0

	StrSql = "Select * From USUARIOS where ID = '" & XUSR & "' "
	Set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	If Not rsX2.eof Then
		XSECTORIG = rsX2(21)
		XSECTORIGB = XSECTORIG
		XSECTGEREN = 0
		'response.write StrSql & "---" & XSECTORIGB
		Do
			'Sector Principal
			StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
			StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco =0" & XSECTORIGB
			'response.write StrSql
			Set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If rsX3.EOF Then
				Exit Do
			Else
				XSECTGEREN = rsX3(0)
				XSECTORIGB = rsX3(0)
				StrSql = "SELECT * FROM CSECOSNJ where ECO = " & XSECTORIGB
				Set rsX4 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
				If rsX4.EOF Then
					Exit Do
				Else
					XJERAR = RSX4(3)
					XJERAR = InStr(XJERAR,"GEREN")
					'anula el control de escalar en el sector
					'  IF XJERAR > 0 THEN
					Exit Do

					'  END IF
				End If
			End If

		Loop



	End If
	If Trim(XSECTORIG & "") = "" Then
		SESSION("SECTORIG")=0
	Else

		SESSION("SECTORIG") = XSECTORIG
	End If


	If Trim(XSECTGEREN & "") = "" Then
		SESSION("SECTGEREN")=0
	Else
		SESSION("SECTGEREN") = XSECTGEREN
	End If



End Function

Function GETDH()

	XKD = DatePart("d",Date())
	XKM = DatePart("m",Date())
	XKY = DatePart("yyyy",Date())





	GETDH = XKD & "/" & XKM & "/" & XKY & " " & Time()


End Function

Function GetProceso(xcodigo)


	xdescri =""
	sinput=""
	Strsql = ""
	Strsql = Strsql & " select x.periodo, A.*, W.ESTADO "
	Strsql = Strsql & " from csrelpathdescri a   "
	Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "
	Strsql = Strsql & " inner join orrskcircuito w on (a.cobj = w.codobjeto)   "
	Strsql = Strsql & " inner join orperiodos x on (w.periodo = x.objid )  "
	Strsql = Strsql & " where  w.objid = 0" & xcodigo & "   "

	'
	'     response.write strsql & chr(13)


	Set rsX1 = ofv.crearconsultaEx(Strsql, cn3, 1, parametros)
	If Not rsX1.EOF Then

		For WKI = 1 To 8

			WKIPROC = rsX1(WKI) & ""

			If WKIPROC <> "" Then
				If WKIPROCD = "" Then WKIPROC =  WKIPROC
				WKIPROCD = WKIPROCD & WKIPROC & " / "
			End If


		Next

		If Len(WKIPROCD) > 4 Then
			WKIPROCD = Left(WKIPROCD, Len(WKIPROCD) - 3)
		End If


		sInput = WKIPROCD & " - (" & rsX1(0) & ")"

	End If
	Call ofv.cerrarconsulta(rsX1)
	getproceso=sinput
End Function

%>