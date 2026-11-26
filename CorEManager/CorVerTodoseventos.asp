<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	Dim wstipo(6)
	aplcheck()

	nCodigo = request.querystring("ncodigo")

	If ncodigo="" Then
		ncodigo=request.querystring("Codobjeto")
	End If

	'response.write ncodigo & "<br>"
	sTipo = request.querystring("Tipobjeto")
	sSectorx = request.querystring("Sector")
	' response.write sSectorx & " 222222224444"
	sMenux = request.querystring("Menu")
	sperfil=session("sperfil")
	addperfil=""
	Set cn = ofv.conectar(ofv.strconn2)
	If smenux="PORSECTOR" Then
		If sperfil="nsnc" Or sperfil="" Then
			addperfil=""
		Else
			Set rspe = ofv.crearconsultaEx("select descri from tipoperfil where tipoperfil='" & sperfil & "'",cn,1,parametros)
			If Not rspe.eof Then
				addperfil= " and  perfdescri='" & rspe(0) & "' "

				Call ofv.cerrarconsulta(rspe)
			End If
		End If
	End If

	If sMenux = "PORSECTOR" Then
		sSelecx = " and ECOdescri= '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "' " & addperfil
	Else
		sSelecx = ""
	End If




	'/----------------------------PDC -----------------------------------------------\


	If  session("APLMDCLI") = "BNX" Or session("APLMDCLI") = "BC" Then


		If CLng("0" & REQUEST.FORM("sector")) <> 0  Then
			SECTORF=REQUEST.FORM("sector")
			fsector = "codsector=" & SECTORF & " and "
			Perfilf="TODOS"
			Cperfil=""
		Else
			SECTORF=0
			fsector=""
			Perfilf="TODOS"
			Cperfil=""
		End If



		If CStr(request.querystring("SECTORF")) <> CStr(SECTORF)  Then
			Perfilf="TODOS"
			Cperfil=""

		Else

			Perfilf=REQUEST.FORM("perfil")

			If perfilf ="TODOS" Then

				Cperfil=""

			Else

				Cperfil = "codperfil='" & Perfilf & "' and "

			End If

		End If



		'SECTOR

		sql = "SELECT distinct b.descri, b.eco"
		sql = sql & " from objsector a inner join eco b on (b.eco=a.codsector)"
		sql = sql & " where a.codobjeto = "  & nCodigo




		Set rsSectorn = ofv.crearconsultaEx(sql, cn, 1, parametros)

		scomK = "0"'"null" & Chr(9)
		scomD = "TODOS" '"" & Chr(9)
		SCOMSEP = Chr(9)
		Do While Not rsSectorn.eof
			scomK=scomK & SCOMSEP & rsSectorn(1)
			scomd=scomd & SCOMSEP & rsSectorn(0)
			SCOMSEP = Chr(9)
			rsSectorn.movenext
		Loop



		sInput =  ofv.gencomboFX("filtrosector","sector",SECTORF,scomk,scomD,"") '& "gencombo"

		Call ofv.cerrarconsulta(rsSectorn)

		celda= celda & ofv.GenCelda("","tu","","5%","","","",  "SECTOR: " ,"si")
		Celda = Celda & ofv.GenCelda("","tu","","45%","","","",sInput,"si")


		'PERFILES


		sql = " select distinct c.descri,c.tipoperfil "

		sql = sql & " from  "
		sql = sql &" OBJPROCED a "
		sql = sql & " inner join objsector b on ( b.objid=a.codsector)"
		sql = sql & " inner join  tipoperfil c on (c.tipoperfil=a.codperfil)"
		sql = sql & " where a.codobjeto=" & nCodigo &"  and  b.codsector=" & SECTORF




		Set rsSectorn = ofv.crearconsultaEx(sql, cn, 1, parametros)

		scomK = "TODOS"'"null" & Chr(9)
		scomD = "TODOS" '"" & Chr(9)
		SCOMSEP = Chr(9)
		Do While Not rsSectorn.eof
			scomK=scomK & SCOMSEP & rsSectorn(1)
			scomd=scomd & SCOMSEP & rsSectorn(0)
			SCOMSEP = Chr(9)
			rsSectorn.movenext
		Loop

		If SECTORF=0 Then
			sinput=""
			celda= celda & ofv.GenCelda("","tu","","5%","","","",  "&nbsp;" ,"si")
		Else

			sInput = ofv.gencomboFX("filtrosector","perfil",Perfilf,scomk,scomD,"") '& "gencombo"
			celda= celda & ofv.GenCelda("","tu","","5%","","","",  "PERFIL: " ,"si")
		End If

		Call ofv.cerrarconsulta(rsSectorn)

		Celda = Celda & ofv.GenCelda("","tu","","45%","","","",sInput,"si")

		naccion="reportesgen.asp?ID=" & nCodigo & "&Titulo=Esquemas_Normativos&tipobj="&sTipo&"&Procedfiltrado=S&Sectorf=" &  SECTORF  & "&perfilF=" & Perfilf

		Call ofv.cerrarconsulta(rsSectorn)


		sInput= "<a target=_new href='" & nAccion & "'>IMPRIMIR Filtrados</b></a>"

		Celda = Celda & ofv.GenCelda("","bietx","","","","","",sInput,"si")

		fila = ofv.GenRow("","","","","","","",celda,"si")


		tabla= ofv.GenTabla("","","","100%","","0","","-1","2","","",fila,"si")

		sAccion="CORVERTODOSEVENTOS.ASP?ncodigo=" & nCodigo & "&Tipobjeto=" & sTipo & "&Sector=" & sSectorx    & "&Menu=" & sMenux   & "&sperfil=" & sperfil & "&SECTORF=" & SECTORF

		Formulariofiltro =   ofv.GenForm("filtrosector","",sAccion,"","post","",tabla,"si")




	Else

		Formulariofiltro =""



	End If

	'/----------------------------PDC -----------------------------------------------\


	Strsql = "SELECT * "
	Strsql = Strsql & " from CSCEventoSectorPerfil Where " & fsector & Cperfil & " obcodobjeto = " & nCodigo
	Strsql = Strsql & sSelecx & " order by secuen, secuencia,codproced "

	CantRegAMostrar= 1
	CantRegAMostrar2= 4
	If BrowserName() = "IE" Or BrowserName() = "CHROME" Then CantRegAMostrar2= 5

	'' sHTML = ofv.FormBody2

	sHTML = ofv.FormHeader(session("FormN"))  & Chr(13)
	shtml= shtml &  "<script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & Chr(13)
	shtml = shtml & CorPnlFnc("N","")


	sMenu = Session("menu")
	bModificar = session("Modificar")
	vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
	vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
	vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)



	sHTML = sHTML  & "<DIV ID='DCEX1' style='display:block;' >"





	' Detalle del procedimiento



	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"  &"<div id=zglos class='norma-card'>"  & Formulariofiltro




	strsqlx = "Select Objprty.Version ,objetos.objun From Objprty  "
	strsqlx = strsqlx & " inner join objetos on(objetos.codobjeto=Objprty.codobjeto) "
	strsqlx = strsqlx & " Where Objprty.codobjeto = " & nCodigo

	Set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
	bVersion = "NO"
	Cobjun=1
	If Not rs.eof Then
	     Cobjun=rs(1)
		If rs(0) = "V" Then bVersion = "SI"
	End If
	Call ofv.cerrarconsulta(rs)
	'response.write strsql
   ' armo sectores
   	if Cobjun=1 and Session("EMPMUN") = "S" then
   	  sql="exec SpSectoresCorp " & ncodigo &",'" & sTipo & "'," & session("un")

   	  Set rsSecCorp = ofv.crearconsultaEx(sql,cn,1,parametros)

           if not rsSecCorp.eof then

               ASecCorp =rsSecCorp.getrows
            end if
          Call ofv.cerrarconsulta(rsSecCorp)
    end if

	Set rsdd = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	eveant  = ""

	EVANT = 0


	Do While Not rsdd.eof
		nID   = rsdd(0)
		tit   = rsdd(7)
		frec  = rsdd(8)
		EVSC  = rsdd(2)
		EVPS  = rsdd(6)
		EVDS  = rsdd(4)

		if isarray(ASecCorp) and Cobjun=1 and Session("EMPMUN") = "S" then
		 SECD  = rsdd(3)
		 SECD = FindInArray(ASecCorp,6,rsdd(9),1,SECD)

		else
		 SECD  = rsdd(3)
		end if
		PERD  = rsdd(5)
		sEven = ofv.convertircar(sEven," ","_")

		Fila = ""

		sPoint = "EV" & nID

		'''''''''' arma tabla de indices '''''''''''''''''''''''''
		stexto1 = tit

		Celdat  = stexto1

		stexto1 = "<A class=nt href='javascript:void(0);' "
		stexto1 = stexto1 & "   >" & Celdat & "</a>"

		Celdat  = "<td class=uts align=center >" & sfidk0 & "</td>"
		Celdat  = Celdat & "<td class=utg style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
		'  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")



		If eveant <> rsdd(4) Then



			Celdat  = "<td class=uts >" & sfidkA & "</td>"
			Celdat  = Celdat & "<td class=utg  >" & rsdd(2) & ".&nbsp;" & rsdd(4) & "</td>"
			Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

			eveant  = rsdd(4)
		End If


		sfidkD  = sfidkC
		stexto1 = rsdd(6) & ".&nbsp;" & rsdd(3) & "&nbsp;-&nbsp;" & rsdd(5)


		Celdat  = stexto1

		stexto1 = "<A class=nt href='javascript:void(0);' "
		stexto1 = stexto1 & "   >" & Celdat & "</a>"


		''   stexto1 = "<A ID='ASC" & rsdd(0) & "' NAME='ASC" & rsdd(0) & "' class=nt href='" & sAccion & "' " & sOnClik1 & " "
		''   stexto1 = stexto1 & "   >" & Celdat & "</a>"

		Celdat  = "<td class=utis >" & sfidkD & "</td>"
		Celdat  = Celdat & "<td class=uti style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
		Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")








		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		If EVANT <> EVSC Then

			If EVANT <> 0 Then
				sHTML = sHTML & "<br>"
			End If

			EVANT = EVSC
			Celda = ofv.GenCelda(sPoint,"ttt","","10%","","","","EVENTO:","si")

			stexto1 = EVSC & "." & EVPS & "&nbsp;-&nbsp;" & EVDS

			Celda = Celda & ofv.GenCelda("","ttt colspan=3 ","","60%","","","",stexto1,"si")

			Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")


		End If


		Celda = ofv.GenCelda("","ut","","10%","","","","SECTOR:","si")

		Celda = Celda & ofv.GenCelda("","vti","","60%","","","",SECD,"si")



		Celda = Celda & ofv.GenCelda("","ut","","10%","","","","FRECUENCIA:","si")

		Strsql = "SELECT descri from frecuencia where codfrec = '" & frec & "' "
		Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs1.eof Then
			wsfrec = rs1(0)
		Else
			wsfrec = "SIN&nbsp;DEFINIR"
		End If

		wsfrec = wsfrec

		Celda = Celda & ofv.GenCelda("","vti","","20%","","","",wsfrec,"si")


		Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")




		Celda = ofv.GenCelda("","ut","","10%","","","","PERFIL&nbsp;:","si")


		Celda = Celda & ofv.GenCelda("","vti colspan=3 ","","60%","","","",perD,"si")


		Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")



		Strsql = "SELECT referencia from objproced where codproced = '" & rsdd(0) & "' "
		Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs1.eof Then
			wsref = rs1(0)
		Else
			wsref = "NO"
		End If

		If IsNull(wsref) Then wsref = "NO"


		' Referencia del procedimiento
		If wsref <> "NO" Then
			sInput = ""
			sInput = sInput  & Replace(wsref,Chr(13),"<br>") & "<br><br>"
			Celda = ofv.GenCelda("","utg style='text-align:justify;' ","","100% colspan=4","","","",sInput,"si")
			Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
		End If

		sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","100% border=0 ","","0","0","-1","2","","",Fila,"si")

		'  response.write sHTML
		' sHTML = ""
		Fila = ""

		bColor = "False"
		If nID <> "no" Then

			ReDim wsmodo(0)
			ReDim wsmodoB(0)

			ym = 1

			strsql = "select * from modo "
			strsql = strsql & " Order By tipomodo"
			Set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs2.eof Then
				Do Until rs2.eof
					tm = ym
					ReDim Preserve wsmodo(ym)
					ReDim Preserve wsmodoB(ym)
					wsmodo(ym) = rs2(0)
					wstipo(0) = rs2(2)
					For ymm = 1 To 6
						wstipo(ymm) = rs2(ymm + 3)
					Next
					wsmodoB(ym) = wstipo
					ym = ym + 1
					rs2.movenext
				Loop
			Else
				tm = ym
				ReDim Preserve wsmodo(ym)
				ReDim Preserve wsmodoB(ym)
				wsmodo(ym) = "NO"
			End If





			strsql = "select * from objaccion "
			strsql = strsql & "WHERE codproced = '" &  nid & "'"
			strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
			strsql = strsql & " Order By secuen"
			' response.write strsql & " 222222 <br>"
			Set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			posicion = 0

			Do While Not rs2.eof

				nAccion   = rs2(0)
				nAccionID = rs2(18)
				xfclas    = "dt"
				sImage    = "&nbsp;"

				If bVersion = "SI" And rs2(14) And SESSION("SHWMOD") <> "N" Then
					If rs2(15) = "M" Then
						xftitl = "Modificado."
						xfclas = "btbs"
						sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
						sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
					ElseIf rs2(15) = "A" Then
						xftitl = "Agregado."
						xfclas = "btba"
						sglmod = "<div class='" & xfclas & "' ><font class=btic >6</font></div>"
						sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
					ElseIf rs2(15) = "E" Then
						xftitl = "Eliminado."
						xfclas = "btbn"
						sglmod = "<div class='" & xfclas & "' ><font class=btic >5</font></div>"
						sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
					End If
				End If



				sPoint = "T" & nAccion


				'Acciones
				If rs2(4) <> "" Then
					Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & rs2(4)
					Set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If  Not rsAcciones.eof Then
						sAccionM = rsAcciones(0)
						sModo   = rsAcciones(1)
					End If
					Call ofv.cerrarconsulta(rsAcciones)
				End If


				Celda  = ofv.GenCelda(sPoint,"uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")


				tenc = "NO"
				For ym=1 To tm
					If ucase(sModo) = ucase(wsmodo(ym)) Then
						tenc = "SI"
						For ymm=0 To 6
							wstipo(ymm) = wsmodoB(ym)(ymm)
						Next

						Exit For
					ElseIf ucase(sModo) < ucase(wsmodo(ym)) Then
						tenc = "NO"
						Exit For
					End If
				Next

				If ym > tm Then  tenc = "NO"

				If   tenc = "NO" Then
					For ymm=0 To 6
						wstipo(ymm) = 0
					Next
				End If




				sLinkalt = "&nbsp;"
				sLink = "&nbsp;"
				''''      if rs2(5) = "CN" then
				If wstipo(1) Then
					'Acciones alternativas
					CantRegAMover = "A"
					vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)
					sID = "Accion=" & nAccionID & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
					sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
					sID = sID & "_Tipobjeto=" & sTipo & "_proced=" & nID & "_Menu=" & sMenux & "_Sector="
					sID = sID & ofv.convertircar(sSectorx,"_","@")

					slinkalt=acciones(nAccionID,nCodigo,nID,sTipo)

				End If


				'Secuencia
				'     Celda = Celda & ofv.GenCelda("","dt","","4%","","","",rs2(10),"si")



				sVerbo   = sAccionM
				sAcciont = sAccionM
				i        = InStr(sVerbo," ")

				If i > 0 Then
					sVerbo   = Left(sAccionM,i-1)
					sAcciont = Mid(sAccionM,i+1)
				Else
					sAcciont=""
				End If

				sInput = rs2(10) & ".&nbsp;<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
				sInput = sInput & sAcciont
				if rs2(4)=301 then
				 Celda  = Celda & ofv.GenCelda("","qt colspan=5  ","","","","","","<b>" & rs2(10) &  ".&nbsp;" & rs2("descriacc") & "</b>","si")
				else
				 Celda  = Celda & ofv.GenCelda("","qt colspan=5  ","","","","","","<b>" & sInput & "</b>","si")
				end if
				Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

				'''''''''' arma tabla de indices '''''''''''''''''''''''''
				stexto1 = sInput
				Celdat  = stexto1

				stexto1 = "<A class=nt href='javascript:void(0);' "
				stexto1 = stexto1 & "   >" & Celdat & "</a>"

				Celdat  = "<td class=utis align=right >" & sfidkB & "</td>"
				Celdat  = Celdat & "<td class=utgi style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
				' Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''








				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				sImagen  = "NO"
				wasotxt  = "N"


				'link
				If wstipo(3) Then
					wlnkptr = "0" & rs2(19)
					If Not IsNull(wlnkptr) Then
						If CLng(wlnkptr) > 0 Then

							Strsql = "SELECT * From objlnk where objid = " & rs2(19)
							Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If Not rsx.eof Then


								wasotxt     = rsx(14) & ""
								wasotxtstr  = CLng("0" & rsx(15))
								wasotxtcan  = CLng("0" & rsx(16))
								wasotxtdat  = rsx(17) & ""

								If wasotxt = "" Then wasotxt = "N"


								wtipoL   = rsx(6)
								wtipoLc  = rsx(7) & ""
								wUrl     = rsx(8) & ""
								wUrlDesc = rsx(9) & ""
								wtipoLNK = rsx(10) & ""

								wNV1    = rsX(11) & ""
								wNV2    = rsX(12) & ""
								wNV3    = rsX(13) & ""

								If  WNV1 = "" Then WNV1 = "0"
								If  WNV2 = "" Then WNV2 = "0"
								If  WNV3 = "" Then WNV3 = "0"


							Else
								wtipoL   = 0
								wtipoLc  = "NO"
								wUrl     = "NO"
								wUrlDesc = ""
								wtipoLNK = "NO"

								WNV1 = "0"
								WNV2 = "0"
								WNV3 = "0"

							End If
							Call ofv.cerrarconsulta(rsx)

							If wtipoLNK = "I" Then

								wNomLt    = wUrlDesc
								wNomL     = wUrlDesc
								If wNomL <> "" Then
									wNomL = UCase(wNomLt)
								Else
									wNomL = "S/D"
								End If

								If Not IsNull(wUrl) Then
									sHref = InStr(wUrl,"://")
									If sHref > 0 Then
										sHref = wUrl
									Else
										sHref = TipoSite("")  & wUrl
									End If
								Else
									sHref = "javascript:void(0);"
								End If
								sFuncion = " target=_blank "


							Else

								wNomLt    = wUrlDesc
								wNomL     = wUrlDesc
								If wNomL <> "" Then
									wNomL = UCase(wNomLt)
								Else
									wNomL = "S/D"


									Strsql = "Select * From csGenObjetos where codobjeto = " & wtipoL
									Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
									If Not rsx.eof Then
										wNomL = rsx(1)
									Else
										wNomL = "S/D"
									End If
									Call ofv.cerrarconsulta(rsx)

								End If

								sActcon = "Consultar.asp?Codigo=" & wtipoL & "&Tipo=" & wtipoLc
								sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
								sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

								sHref = "&nbsp;"

								sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

							End If


							slnktit = wNomLt

							SArma = wNomL

							Slnkdata = wNomL

							sImagen  = "SI"

							sLink       = ofv.GenLink("","btbn title='" & slnktit & "' ",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),sgllnk,"si")
							sLink2       = ofv.GenLink("","nt title='" & slnktit & "' ",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),"NNNNNN","si")



						End If
					End If
				End If




				''''''''''''''''''''''''''''''''''''''


				'Texto de Referencia
				sArma = rs2(11)
				If Len(sArma) > 0 Then


					If wasotxt = "S" Then
						If  wasotxtstr > 0  And wasotxtcan > 0 Then

							sfontstyl = " style='font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("APLFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:800;' "

							sArma = Left(sArma,wasotxtstr - 1) & "" & Replace(sLink2,"NNNNNN","&nbsp;<u><font " & sfontstyl & "  >" & Mid(sArma,wasotxtstr,wasotxtcan) & "</font></u>&nbsp;") & "" & Right(sArma,Len(sArma) - (wasotxtstr + wasotxtcan - 1))
						End If
					End If

					sArma = Replace(sArma,Chr(13),"<br>")
					Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
					Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
				End If

				sArma = rs2(16)
				If Len(sArma) > 0 Then
					sArma = Replace(sArma,Chr(13),"<br>")
					Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
					Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
				End If

				sArma = rs2(17)
				If Len(sArma) > 0 Then
					sArma = Replace(sArma,Chr(13),"<br>")
					Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
					Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
				End If


				'link
				If wstipo(3) Then

					If  wasotxt  = "N" Then
						If  simagen <> "NO" Then

							sLink2 = Replace(sLink2,"NNNNNN",slnkdata)

							Celdak = ""
							Celdak = Celdak & ofv.GenCelda("","btbn title='ver documento.' ","","2","","","",sLink,"si")
							Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sLink2,"si")

							Filak = ofv.GenRow("","","","","","","",Celdak,"si")
							SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")


							Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
							Celda = Celda & ofv.GenCelda("","dt ","","","","","","&nbsp;&nbsp;","si")
							Celda = Celda & ofv.GenCelda("","dt2 colspan=3","","","","","",SArma,"si")
							Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

						End If

					End If


				End If


				''''''''''''''''''''''''''''''''''''''''''''''''''''''''






				If rs2(6) <> "" And rs2(6) > "0"  Then
					Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")

					nCopias = "&nbsp;"

					If wstipo(2) Then
						'Copias
						If Not IsNull(rs2(13)) Then
							If rs2(13) > "0" Then
								If rs2(13) > "1" Then
									nCopias = "&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copias"
								Else
									nCopias = "&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copia"
								End If
							End If
						End If
					End If

					'Componentes
					sComponente = "&nbsp;"
					sLink       = sglcmp
					sLink2      = "&nbsp;"

					Strsql = "Select codobjetol,tipobjetol,* From objlnk Where ltipo  in('KL','KF')  and objid = 0" & rs2(6)
					Set rsCompoM = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If  Not rsCompoM.eof Then
						wNV1    = rsCompoM(13) & ""
						wNV2    = rsCompoM(14) & ""
						wNV3    = rsCompoM(15) & ""

						If  WNV1 = "" Then WNV1 = "0"
						If  WNV2 = "" Then WNV2 = "0"
						If  WNV3 = "" Then WNV3 = "0"



					 STIPOL=rsCompoM("LTIPO")

					   SELECT CASE STIPOL

                          CASE "KF"
                           codobjetoL = clng(rsCompoM("codobjetoL"))
						   wtipoLc    = rsCompoM(9)
                           wUrlDesc = rsCompoM(11) & ""


							wNomL    = wUrlDesc
							If wNomL = "" Then
								wNomL = "S/D"


								Strsql = "Select * From csGenObjetos where estado <> 4 and codobjeto = " & codobjetoL
								'      RESPONSE.WRITE STRSQL
								Set rst2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								If Not rst2.eof Then
									wNomL = rst2(1)
								Else
									wNomL =""' "S/D"
								End If
								Call ofv.cerrarconsulta(rst2)

							End If

							sActcon = "Consultar.asp?Codigo=" & codobjetoL & "&Tipo=" & wtipoLc
							sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
							sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3



							sHref    = "&nbsp;"
							' response.write sActcon
							sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;" & Chr(34)


							slnktit  = wNomL
							wNomL    = pStyle & "&nbsp;<u><i>" & wNomL & "</i></u>&nbsp;" & pStyleF
							sImagen  = wNomL
							sLink       = ofv.GenLink("","btba",GenKeyLnks(sActcon),"_new","","",sglcmp,"si")
							sLink2      = ofv.GenLink("","nt",GenKeyLnks(sActcon),"_new","","",wNomL,"si")

                        CASE "KL"
                          Strsql = "Select nombre From csacomple Where codcomp = " & rsCompoM(0)
						   Set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						  If  Not rsComponentes.eof Then
							sComponente = rsComponentes(0)
							SArma       = Replace(sComponente," ","&nbsp;")  & nCopias & ""
							sHref       = "EditarImagen2.asp?ID=" & rsCompoM(0)
							sHref       = sHref & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

							sLink       = ofv.GenLink("","btba",GenKeyLnks(sHref),"_new","","",sglcmp,"si")
							sLink2      = ofv.GenLink("","nt",GenKeyLnks(sHref),"_new","","",SArma,"si")

						  End If
						Call ofv.cerrarconsulta(rsComponentes)
                       end select





					End If
					Call ofv.cerrarconsulta(rsCompoM)

					Celdak = ""
					Celdak = Celdak & ofv.GenCelda("","btba style='cursor:pointer;' title='ver complemento.' ","","2","","","",sLink,"si")
					Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sLink2,"si")

					Filak = ofv.GenRow("","","","","","","",Celdak,"si")
					SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")


					Celda = Celda & ofv.GenCelda("","dt ","","","","","","&nbsp;&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","dt2 colspan=3","","","","","",SArma,"si")
					Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
				End If


				If rs2(8) > "0" Or rs2(9) > "0" Then
					Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
					'Sector Continuador
					sContinuador = "&nbsp;"
					If rs2(8) > "0" Then
						Strsql = "Select descri From Eco Where eco = " & rs2(8)
						Set rsContinuadores = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not rsContinuadores.eof Then sContinuador = rsContinuadores(0) & "&nbsp;"
						Call ofv.cerrarconsulta(rsContinuadores)
					End If


					'Perfil Continuador
					sPerfil = "&nbsp;"
					If rs2(9) > "0" Then
						Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs2(9) & "'"
						Set rsPerfiles = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not rsPerfiles.eof Then sPerfil = rsPerfiles(0) & "&nbsp;"
						Call ofv.cerrarconsulta(rsPerfiles)
					End If


					Celdak = ""
					Celdak = Celdak & ofv.GenCelda("","ut title='sector continuador.' ","","2","","","",sglcnt,"si")
					Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sContinuador & "&nbsp;&nbsp;","si")
					Celdak = Celdak & ofv.GenCelda("","ut title='perfil continuador.'  ","","2","","","",sglprf,"si")
					Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sPerfil,"si")

					''style='border-style:solid;border-width:1;border-color:#eeeeee;'

					Filak = ofv.GenRow("","","","","","","",Celdak,"si")
					SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")

					Celda = Celda & ofv.GenCelda("","dt ","","","","","","&nbsp;&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","dt colspan=3","","","","","",SArma,"si")

					Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
				End If

				' if  wstipo(1) then
				Celda =  ofv.GenCelda("","dt colspan=2","","1%","","","","&nbsp;","si")
				Celda =  celda & ofv.GenCelda("","dt colspan=5","","","","","",sLinkalt,"si")
				Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
				If sLinkalt <> "" And   sLinkalt <> "&nbsp;"   Then
					' if sLinkalt <> "" and   sLinkalt <> "&nbsp;"   then

					If session("APLMDCLI")="BC" Then
						Celda =  ofv.GenCelda("","vt colspan=2","","1%","","","","&nbsp;","si")
					Else
						Celda =  ofv.GenCelda("","vt colspan=2","","1%","","","","SI","si")
					End If
					Celda =  celda & ofv.GenCelda("","dt colspan=5","","","","","","&nbsp;","si")
					Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
				End If
				sLinkalt = ""
				'end if


				Celda    = ""

				Celda = Celda & ofv.GenCelda("","dt","","1%","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","dt","","1%","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","dt","","3%","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","dt","","50%","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")

				Fila = Fila & ofv.GenRow("","","","","5","","",Celda,"si")

				rs2.movenext
				posicion = posicion + 1
			Loop
		End If
		sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;'","","95% border=0 ","","0","0","0","0","","",Fila,"si")
		' next
		rsdd.movenext
	Loop

	sHTML = sHTML & "<br>"

	If nID <> "no" Then Call ofv.cerrarconsulta(rs2)

	Call ofv.cerrarconsulta(rs)



	'''  sHTML = sHTML & ofv.GenTabla("","","","95%","","0","0","0","0","","",Fila,"si")

	sHTML = sHTML & "</div>"

	sHTML = sHTML & "</div>"




	Celda = ofv.GenCelda("","ut","","100% colspan=2","","","","&nbsp;","si")
	Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	'sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")



	''''''' selector del glosario

	GLhtml = VerGlosario(shtml)

	sHTML = GLhtml

	Call ofv.cerrarconn(cn)

	''''''''''''''''''''''''''''''



	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")

	'Botonera
	Celt1 = ""
	'' Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
	Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")

	Agregar = Celt2
	Buscar  = Celt1
	Celda   = ""

	If bVersion = "SI" Then
		If SESSION("VIEWMOD") <> "N" Then

			'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
			Celda   = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

			sasp    = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
			sHTML   = sHTML & ConsCambios(sasp)


		End If


	End If


	sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


	sHTML = sHTML & GTCOBS() & "</div></body></html>"
End If

response.write sHTML

%>

<%

Function acciones(nAccion,nCodigo,nProced,sTipo)
	Dim shtml
	Dim rs,rs2
	Set cn = ofv.conectar(ofv.strconn2)

	strsqlx = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
	Set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
	bVersion = "NO"
	If Not rs.eof Then
		If rs(0) = "V" Then bVersion = "SI"
	End If
	Call ofv.cerrarconsulta(rs)


	strsql = "SELECT TIPOACCALT.tipoaccionalt, TIPOACCALT.descri, Objaccion.secuen "
	strsql = strsql & "FROM OBJACCION, TIPOACCALT where "
	strsql = strsql & "OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
	strsql = strsql & "and OBJACCION.codaccion = '" & nAccion & "'"
	strsql = strsql & " Order By secuen"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then
		VTACC = rs(0)
		VDACC = rs(1)
		VSACC = rs(2)
	End If
	Call ofv.cerrarconsulta(rs)




	strsql = "SELECT * from csaccaltsel "
	strsql = strsql & "Where idaccion = '" & nAccion & "' "
	strsql = strsql & " and (modi = 0 or "
	strsql = strsql & " (modi <> 0 and tipmodi <> 'E')) "
	strsql = strsql & " order by secuencia"
	'response.write strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)



	strsql = "SELECT * FROM TIPOACCALT WHERE TIPOACCALT.modo='OP' Or TIPOACCALT.modo='RF' "
	strsql = strsql & "ORDER BY descri"


	sHTML = ""

	sHTML = sHTML & ofv.GenTabla("panel","","","80% ","","0","0","-1","2","","","","no")

	t2    = ofv.GenCelda("","utis colspan=3 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ","","","","","",sfidkC & "&nbsp;NO.","si")

	sHTML = sHTML & ofv.GenRow("","","","","","","",t2,"si")


	Vcont = 1
	Valor = ""
	VCTACC = ""
	VCDACC = ""
	VCSACC = ""

	If Not rs.eof Then
		Do Until rs.eof
			Valor = rs(0)
			If Not IsNull(rs(8)) Then
				Continuador = rs(8)


				strsql = "SELECT TIPOACCALT.tipoaccionalt, TIPOACCALT.descri, Objaccion.secuen "
				strsql = strsql & "FROM OBJACCION, TIPOACCALT where "
				strsql = strsql & "OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
				strsql = strsql & "and OBJACCION.codaccion = '" & Continuador & "'"
				strsql = strsql & " Order By secuen"
				'response.write strsql
				Set rs2 = ofv.crearconsultaEx(StrSql,cn,1,volver)
				If Not rs2.eof Then
					VCTACC = rs2(0)
					VCDACC = rs2(1)
					VCSACC = rs2(2)
				End If
				Call ofv.cerrarconsulta(rs2)

			End If
			ID = rs(2)
			sImage = "&nbsp;"
			If bVersion = "SI" And rs(9) And SESSION("SHWMOD") <> "N" Then
				If rs(10) = "M" Then
					sImage = ofv.GenImage("","","","","Imagenes/Botmodx.jpg","0","")
				ElseIf rs(10) = "A" Then
					sImage = ofv.GenImage("","","","","Imagenes/Botagrx.jpg","0","")
				End If
			End If

			t1 = ofv.GenCelda("","dt","","5%","","","",sImage,"si")

			t1 = t1 & ofv.GenCelda("","dt","","5%","","","","&nbsp;" & VSACC & "." & rs(6) & "&nbsp;","si")

			Vcont = Vcont + 1

			'Texto Accion Alt
			If rs(4) <> "" Then
				Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rs(4)
				'response.write strsql
				Set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsAcciones.eof Then
					sAccion = rsAcciones(0)
				End If
				Call ofv.cerrarconsulta(rsAcciones)
			End If

			sVerbo = sAccion
			sAcciont = sAccion
			i = InStr(sVerbo," ")
			If i > 0 Then
				sVerbo = Left(sAccion,i-1)
				sAcciont = Mid(sAccion,i+1)
			Else
				sAcciont=""
			End If
			sInput = "<B><i>" & sVerbo  & "</i>" & "&nbsp;&nbsp;&nbsp;"
			sInput = sInput & sAcciont & "</b>"

			t2 = ofv.GenCelda("","qt ","","90%","","","",sInput,"si")

			sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2,"si")

			'Texto de Referencia
			If BrowserName() = "IE" Or BrowserName() = "CHROME" Then
				nCol = 60
			Else
				nCol = 60
			End If

			t3 = ofv.GenCelda("","","","","","","","&nbsp;","si")
			t3 = t3 & ofv.GenCelda("","","","","","","","&nbsp;","si")
			sArma = rs(7)
			If Len(sArma) > 0 Then
				sArma = Replace(sArma,Chr(13),"<br>")
			End If
			'''    SArma = "<font color=#225588 face=arial-narrow size=1>" & sArma & "</font>"
			t3 = t3 & ofv.GenCelda("","et ","","","","","",SArma,"si")
			sHTML = sHTML & ofv.GenRow("","","","","","","",t3,"si")
			rs.movenext
		Loop


		'Encabezado Continuador



		t4 = ofv.GenCelda("","","","2%","","","","&nbsp;","si")
		t4 = t4 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		'Texto Continuador
		If VCDACC <> "" Then
			sVerbo = VCDACC
			sAcciont = VCDACC
			i = InStr(sVerbo," ")
			If i > 0 Then
				sVerbo = Left(VCDACC,i-1)
				sAcciont = Mid(VCDACC,i+1)
			Else
				sAcciont=""
			End If
			sInput = "<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
			sInput = sInput & sAcciont
			sContinuador = "<b>" & VCSACC & " - " & sInput & "</b>"
		End If
		t4 = t4 & ofv.GenCelda("","et","","","","","","Continua en: " & sContinuador,"si")
		If VCDACC <> "" Then
			sHTML = sHTML & ofv.GenRow("","","","","","","",t4,"si")
		End If

	Else

		t6 = ofv.GenCelda("","","","","","","","&nbsp;","si")
		t6 = t6 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		''   SArma = "<h5><B>Sin&nbsp;datos&nbsp;disponibles</B></h5>"
		SArma = "&nbsp;"
		t6 = t6 & ofv.GenCelda("","dt","","","","","",SArma,"si")
		sHTML = sHTML & ofv.GenRow("","","","","","","",t6,"si")


	End If

	t5 = ofv.GenCelda("","","","","","","","&nbsp;","si")
	t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
	t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
	sHTML = sHTML & ofv.GenRow("","","","","","","",t5,"si")


	sHTML = shtml & ofv.GenFTabla()

	Call ofv.cerrarconsulta(rs)



	acciones=shtml
End Function



%>