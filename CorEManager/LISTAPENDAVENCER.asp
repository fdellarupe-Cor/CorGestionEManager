<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0


	CantRegMostrarx = 8
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

	SPROCESO = REQUEST.QUERYSTRING("PROCESO")
	SFILTRO  = REQUEST.QUERYSTRING("FILTRO")
	SFILSEL  = REQUEST.QUERYSTRING("FILSEL")
	SsubFILTRO  = REQUEST.QUERYSTRING("SUBFILTRO")
	SsubFILTRO2 = REQUEST.QUERYSTRING("SUBFILTRO2")
	SsubFILTRO3 = REQUEST.QUERYSTRING("SUBFILTRO3")

	sMenu = request.querystring("Menu")
	sNombre = request.querystring("Nombre")

	If  SPROCESO<> "" Then
		session("SPROCESO")=SPROCESO
	Else
		SPROCESO=session("SPROCESO")
	End If

	If SsubFILTRO <> "" Then
		session("SsubFILTRO")=SsubFILTRO
	Else
		SsubFILTRO=session("SsubFILTRO")
	End If

	sHTML = ofv.FormHeader(session("FormN"))


	ofv.ObtenerAtributos Session("Form"),""
	bConsultar = ofv.AConsultar
	'response.write request.form
	Set cn = ofv.conectar(ofv.strconn2)



	Fila = ""
	xTrColor = "#F9F9F9"


	sImagevMa    = "<font face='arial' size=1 xcolor='#aa0000' style='font-weight:900;background-color:#aa0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MUY&nbsp;ALTO&nbsp;</font>"
	sImageva    = "<font face='arial' size=1 xcolor='#f7aa22' style='font-weight:900;background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ALTO&nbsp;</font>"
	sImagevm    = "<font face='arial' size=1 xcolor='#f7EEDD' style='font-weight:900;background-color:#f7EEDD;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MODERADO&nbsp;</font>"
	sImagevb    = "<font face='arial' size=1 xcolor='#6699CC' style='font-weight:900;background-color:#6699CC;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ESCASO&nbsp;</font>"
	sImagevNUL    = "<font face='arial' size=1 xcolor='#00AA00' style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NULO&nbsp;</font>"

	sImageSI    = "<font face='arial' size=1  style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;SI&nbsp;</font>"
	sImageNO    = "<font face='arial' size=1  style='font-weight:900;background-color:#AA0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NO&nbsp;</font>"

	sImagevn    = "<font face='arial' size=1 xcolor='#dddddd' style='font-weight:900;background-color:#dddddd;color:#ffffff' >----</font>"

	wkstylof = "AAA style='font-family:arial;font-size:6pt;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' "
	wkstylof2 = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "

	''''b.codvinobj = 'VN' AND

	Set cnX = ofv.conectar(ofv.strconn2)

	Strsql = ""

	'''''Strsql = Strsql & " SELECT *,cast('12/12/2050' as datetime) - (cast(substring(CKTVIEW.segcktpsc,10,5) as integer)  ) as FECCOMIENZO,cast( ( getdate() - ( cast('12/12/2050' as datetime) - (cast(substring(CKTVIEW.segcktpsc,10,5) as integer)  ) )  ) as integer) as DIFER  FROM CKTVIEW "

	dFechadia = CLng(Date) - CLng(CDate("01-01-1900"))
	''response.write dfechadia & "<br>"

		Strsql = Strsql & " select a.autor, a.codobjeto, a.tipobjeto, a.fecvto, a.estado "
		Strsql = Strsql & " from objprty a "
		Strsql = Strsql & " inner join objetos b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto) "
		Strsql = Strsql & " where a.estado in (3)  and b.objun=" & session("UN")
		Strsql = Strsql & " and a.fecvto > '" & dFechadia & "' "
		If CStr(sproceso) <> "0" Then
			''Se agrega a. al query de abajo ya que le faltaba el alias de la tabla''
			Strsql = Strsql & " and a.tipobjeto ='" & sproceso & "' "

		End If




	If CStr(SsubFILTRO) <> "0" Then

		Strsql = Strsql & " and autor ='" & SsubFILTRO & "' "

	End If

	''Strsql = Strsql & " order by autor,codobjeto,fecvto "
	Strsql = Strsql & " order by autor,fecvto,a.codobjeto "

	'response.write strsql & "<br>"

	Set rs = ofv.crearconsultaEx(StrSql,cnX,1,parametros)


	If rs.eof Then

		Celda = ""
		Celda = Celda & ofv.gencelda("","VTG colspan=9 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY DOCUMENTOS A VENCER PARA<BR>LA BUSQUEDA SOLICITADA","si")
		Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


	Else

		posI = 0

		wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;' "


		WKIPROCc = 0
		MAINPROC = 0
		MAINACTI = ""

		Celda = ""
		Celda = Celda & ofv.gencelda("","TT colspan=7 style='color:#ffffff' " & wkstylo1,"center","","","","","LISTA DE DOCUMENTOS A VENCER","si")
		Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

		Celda = ""
		Celda = Celda & ofv.gencelda("","tt " & wkstylo1,"center","","","","","AUTOR/DUEÑO","si")
		Celda = Celda & ofv.gencelda("","tt " & wkstylo1,"center","","","","","TEMA/SUBTEMA","si")
		Celda = Celda & ofv.gencelda("","tt " & wkstylo1,"center","","","","","DOCUMENTO","si")
		Celda = Celda & ofv.gencelda("","tt " & wkstylo1,"center","","","","","AVISO DE ANTELACION","si")
		Celda = Celda & ofv.gencelda("","tt " & wkstylo1,"center","","","","","VENCIMIENTO","si")
		Celda = Celda & ofv.gencelda("","tt " & wkstylo1,"center","","","","","REASIGNAR NUEVO USUARIO/DUEÑO","si")
		Fila  = FILA & ofv.genrow("","aaa bgcolor='#ddddff'  ","","","25","","",Celda,"si")
		'response.write CantRegMoverx
		If Not rs.eof Then rs.move(CantRegMoverx)
		Do While (Not rs.EOF) And POSI < CantRegMostrarx

			smarca = "S"

			Celda = ""

			''' Autor
			sInput = RS(0)
			'''
			''' response.write getInfoUser(RS(0)) & " linea 163 " & chr(13) & "<br>"
			''' response.write session("SGU_nombre") & " linea 164 " & chr(13) & "<br>"
			'''
			xxxx = getInfoUserex(RS(0),session("un"))
			If session("SGU_nombre") <> "" Then
				sInput =  session("SGU_nombre")
			Else
				sInput =  RS(0)
			End If

			Celda = Celda & ofv.gencelda("","UTi style='color:" & Session("CMFPNTITFT") & ";text-align:center;border-style:solid;border-width:1;border-color:#aaaaaa;background-color:color-mix(in srgb, " & Session("CMFPNTIT") & ", transparent 20%);'","","","","","",sInput,"si")

			''`tipobjeto
			spaviso = ""
			stipobj = rs(2) & ""

			Strsql = ""
			Strsql = Strsql & "SELECT * FROM TIPObjeto "
			Strsql = Strsql & "where tipobjeto ='" & stipobj & "' "

			'' agregado 10112014
			Strsql = Strsql & " and avisoant <> 0 "
			'' agregado 10112014

			Strsql = Strsql & "ORDER BY DESCRI"
			'''response.write " tipobj " & strsql & chr(13) & "<br>"
			Set rsxx = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			spaviso = 0

			If Not rsxx.eof Then
				'''sInput = RSXX(1)
				sInput = stipobj
				spaviso = rsxx(24)  ''avisoant
			End If
			Call ofv.cerrarconsulta(rsXX)

			'''
			'''tema/subtema a partir de codobjeto
			sInput = rs(1)
			Strsql = ""
			Strsql = Strsql & " select * "
			Strsql = Strsql & " from csrelpathdescriex a   "
			Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "
			Strsql = Strsql & " where  a.cobj = 0" & sInput & "   "
			''''response.write  chr(13) & Strsql & chr(13)
			Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsx1.eof Then
				sinput=rsx1("temadescri") & " / " & rsx1("stemadescri")
				'''SDOC=RSX1("TOBJDESCRI") & " : " & RSX1("EXPR7")
			Else
				sinput="---"
				'''SDOC=""
			End If

			Call ofv.cerrarconsulta(rsx1)

			Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")

			'''codobjeto
			scodobj = RS(1)
			Strsql = ""

			Strsql = Strsql & " SELECT A.CODOBJETO, (A.DESCRI + ': ' + A.OBJDESCRI) AS NOMBRE "
			Strsql = Strsql & " FROM CSOBJETOSRDO A "
			Strsql = Strsql & " where a.codobjeto = 0" & RS(1)
			Strsql = Strsql & " ORDER BY NOMBRE "
			'response.write " codobj " & strsql & chr(13) & "<br>"
			Set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsX1.eof Then sInput = scodobj & " : " & RSX1(1)
			Call ofv.cerrarconsulta(rsX1)

			Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")

			'''aviso de antelacion

			'' agregado 10112014
			If spaviso <> 0 Then
				sInput = spaviso & " dias "
			End If
			'' agregado 10112014

			'' agregado 6-11-2014

			dfecvto   = rs(3)
			dfecvto   = CLng(dfecvto)
			dFechadia = CLng(Date) - CLng(CDate("01-01-1900"))

			'' RESPONSE.WRITE dFechadia & " dFechadia " & "<BR>"
			'' fin agregado 6-11-2014

			'' agregado 6-11-2014

			'' RESPONSE.WRITE dfecvto & " dfecvto "  & "<BR>"
			'' RESPONSE.WRITE dFechadia & " dFechadia " & "<BR>"
'response.write
			If dfecvto > dFechadia Then
				StrSql = ""
				StrSql = StrSql & "SELECT tipobjeto, avisoant from tipobjeto where aplic in('EM','GLB') "
				StrSql = StrSql & " and tipobjeto = '" & rs(2) & "' "
				''RESPONSE.WRITE STRSQL & "<BR>"
				Set rsy = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				If Not rsy.EOF Then
					'' RESPONSE.WRITE " linea 2190 "  & "<BR>"
					davisoant = 0
					davisoant = rsy(1)
					dfecaviso = rs(3) - davisoant

					''RESPONSE.WRITE davisoant & " davisoant "  & "<BR>"
					''RESPONSE.WRITE dfecvto & " dfecvto "  & "<BR>"

					dfecaviso = CLng(dfecaviso)

					''RESPONSE.WRITE dfecaviso & " dfecaviso "  & "<BR>"
					''RESPONSE.WRITE dFechadia & " dFechadia " & "<BR>"

          If (dfecaviso <= dFechadia and dfecvto >= dFechadia) Then

						diferavis = 0
						diferavis = dfecaviso - dfechadia
						xdife = 0
						xdife = diferavis - davisoant

						''RESPONSE.WRITE diferavis & " diferavis "  & "<BR>"
						''RESPONSE.WRITE xdife & " xdife "  & "<BR>"

						'' fin agregado 6-11-2014
						'' agregado 6-11-2014

						''       End If

						'' fin agregado 6-11-2014

						'' response.write " aviso " & sInput & chr(13) & "<br>"

						Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")

						'''fecha de vencimiento

						fi = CLng(CDate("01-01-1900"))

						''response.write " rs3 " & rs(3) & chr(13) & "<br>"
						If rs(3) > 0 Then
							dfecvto = CDate(CLng(rs(3)) + fi)
							sInput = fechadorCor(dfecvto,"")
							'' response.write sInput & " sInput " & chr(13) & "<br>"


						Else
							sInput = 0
							'''response.write sInput & " sInput x cero"
						End If

						Celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")

						Call ofv.cerrarconsulta(rsxy)

						'''response.write " rs1 " & rs(1) & chr(13) & "<br>"



				'		If rs("estado") = 2 Or rs("estado")=4 Then
							sInput =  "<A HREF='WKFREASIGNADAV.ASP?WKFCOD=" & RS(1) & "&CantRegMoverx=" & CantRegMoverx &"' "
							sInput =  sInput & "CLASS='NT' style='color:" & Session("CMFPNBOT") & "' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>REASIGNAR</U></A>"
				'		Else
				'			sInput =  "<U>REASIGNAR</U>"
							'  sInput =  sInput & "CLASS='NT' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>REASIGNAR</U></A>"
				'		End If

						celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")

						'Celda = Celda & ofv.gencelda("","UTi  " & wkstylo1 & "  ","","","","","",sInput,"si")

						Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

						If xTrColor = "#FFFFFF" Then
							xTrColor = "#F9F9F9"
						Else
							xTrColor = "#FFFFFF"
						End If
					End If

				End If

			End If


			posi = posi + 1
			rs.Movenext
		Loop
		Celda = ""
		Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
		Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
		Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
		Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
		Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
		Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
		Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")


	End If
	sHTML = sHTML & ofv.GenTabla("","aaa ","","95% ","","0","","0","-1","","",Fila,"si")

	'Botonera
	Agregar = ""
	Buscar  = ""

	sAccion = "SUBCONSULTA.asp?Nombre=" & sNombre & "&Menu=" & sMenu
	Celda = ""'ofv.BotonVolver(sAccion)


	'sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
	sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	Call ofv.cerrarconn(cnx)

	sHTML = sHTML & "</body></html>"
End If
Function fechaArg(f)
	dia = Day(f)
	mes = Month(f)
	ano = Year(f)
	fechaArg = dia & "/" & mes & "/" & ano
End Function
Response.Write sHTML

%>