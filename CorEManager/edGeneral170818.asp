<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	XXXX = GENLOGEXT("EDIDET")

	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False

	OPC = request.querystring("OPC")
	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")
	vuelta = ofv.vueltaasp2("../")
	vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

	sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
	sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
	histobs = request.querystring("histobs")

	If histobs = "S" Then

		niv1x=request.querystring("NIV1")

		session("codSol")=niv1x

	End If
	Set cn2 = ofv.conectar(ofv.strconn0)
	Strsqlx = "Select * From NivelConfiden order by NCCOD"
	Set rs = ofv.crearconsultaEx(StrSqlx,cn2,1,parametros)
	ReDim Preserve CNIV(0)
	ReDim Preserve DNIV(0)
	ii = 0
	If Not rs.eof Then
		Do Until rs.eof
			ReDim Preserve CNIV(ii)
			ReDim Preserve DNIV(ii)
			CNIV(ii) = rs(0)
			DNIV(ii) = rs(1)
			ii = ii + 1
			rs.movenext
		Loop

	End If
	Call ofv.cerrarconsulta(rs)

	Set cn = ofv.conectar(ofv.strconn2)

	strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
	''response.write strsql & " strsql "
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	bVersion = "NO"
	If Not rs.eof Then
		If rs(0) = "V" Then bVersion = "SI"
		Nestado=rs(1)
	End If
	Call ofv.cerrarconsulta(rs)

	'response.write OPC & " OPC " & "<br>"

	If OPC = "SI" Then
		ESQ = request.querystring("ESQ")
		''     IF ESQ > 0 THEN
		STRSQL = "UPDATE OBJETOS SET OBJREM = " & ESQ & " WHERE CODOBJETO = " & nCodigo
		' response.write strsql & " strsql " & "<br> "
		Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

		'     sql = "Update novegenpar Set MSGOBJCOD = " & ESQ & "
		'    sql = sql & "' Where MSGOBJCOD = " & nCodigo
		'sql = sql & " and tipobjmaster = '" & xntipobj & "' "
		'     call ofv.crearconsultaEx(Sql,cn,1,volver)


		''     END IF
	End If

	sImagenEM  = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0143;</FONT></B>"
	sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"
	sImagenx   = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

	sHTML = ofv.FormHeader(session("FormN"))



	shtml = shtml & CorPnlFnc("N","")

	'----------------------------------------------


	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span>" ' agregado para chrome y firefox <span>

	'CONTROL ESTADO 4 LA SOLICITUD PARA PODER CARGARLA
	If session("BASE") = "D" Then

		fSql = "select ISNULL(nrosolic,0) from objprty where codobjeto=" & nCodigo & " and tipobjeto='" & sTipo & "' and estado=1"
		Set xs = ofv.crearconsultaEx(fSql,cn,1,parametros)

		If Not xs.eof Then

			If xs(0) > 0 Then

				xSql = "select estado from pedidos where objid=" & xs(0) & " and estado=4"
				Set xd = ofv.crearconsultaEx(xSql,cn,1,parametros)

				If xd.eof Then

					sSql = "update objprty set nrosolic=NULL where codobjeto=" & nCodigo & " and tipobjeto='" & sTipo & "'"
					Call ofv.crearconsultaEx(sSql,cn,1,parametros)

					sHTML = sHTML &  "<script>"
					sHTML = sHTML &  "alert('La solicitud debe estar en estado Asignado y en ejecucion.');"
					sHTML = sHTML &  "</script>"

				End If
				Call ofv.cerrarconsulta(xd)
			End If

		End If
		Call ofv.cerrarconsulta(xs)


	End If


	Strsql = "Select Objetos.*,b.codinx2,objprty.nrosolic,objprty.objid as idZ From Objetos "
	'strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (Objetos.codobjeto = b.codinx2) "
	strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where TIPO='OBS' AND actpend is null and codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on (cast ( Objetos.codobjeto as nvarchar)= b.codinx2) "
	strsql = strsql & " inner join objprty on (objetos.codobjeto=objprty.codobjeto and objetos.tipobjeto=objprty.tipobjeto) "
	strsql = strsql & " Where objetos.codobjeto=" & nCodigo

	'sHTML = sHTML &  "<br>" &  Strsql &  "<br>"
	'response.write  chr(13) & Strsql & chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If  Not rs.eof Then

		SVIENE = "N"

		wdescri = rs(3)

		Fila = ""

		EsCopr = "N"

		'response.write nCodigo & " nCodigo " & "<br>"
		'response.write sTipo   & " sTipo " & "<br>"

		''''' agregado para color del cuadrito de Observaciones version 4.6
		wComentario = ""
		strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and "
		strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"      ' and "
		strsql = strsql & " order by codcircuito, codusr, objid"
		''response.write strsql
		Set rss = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rss.eof Then
			wComentario = rss(13)
		End If
		Call ofv.cerrarconsulta(rss)

		If (Nestado=5  Or Nestado=6) And (rs(20) & "") <> "" And bModificar Then

			sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
			sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(0)  & "&ID2=" & rs(0)

			xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
			sHref   = sAccion

			''''''agregado para color del cuadrito de Observaciones version 4.6

			If wComentario = "" Then
				sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
				'''''''sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"

			Else
				sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
			End If
			''''''



			'se cambio link por celda'
			Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")

			Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")



			'Codigo
			Celda  = ofv.GenCelda("","tt","","20%","","","","Observaciones","si")
			Celda = Celda & Celdaobs
			Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

		End If






		If  Session("EMPMUN") = "S" Then

			'empresa
			Celda = ofv.GenCelda("","tt","","","","","","Empresa","si")

			Strsql  = "Select * From clienteplan where id = " & rs(7)
			Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  Not rsx.eof Then
				sInput  = rsx(1)
				sInput2 = rsx(5)
			End If


			sinput   = "<img src='../COREss20/Empresa/" & sInput2 & "' >" & "&nbsp;" & sinput

			SINPUT2  = ""

			If CLng("0" & rs(7)) = 1 Then EsCopr = "S"

			Celda = Celda & ofv.GenCelda("","uti valign=top ","","","","","",sInput & SINPUT2,"si")
			Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

		End If

		'Codigo
		Celda  = ofv.GenCelda("","tt","","20%","","","","Codigo","si")
		sInput = nCodigo & sTipo & "/" & RS("aplic")
		Celda  = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

		'Orden
		Celda = ofv.GenCelda("","tt","","20%","","","","Orden","si")
		If  bModificar Then
			'sInput = ofv.GenerarInput(rs(13).name,rs(13),"text","edGeneral","1","dtn","")
			sInput = "<input id=" & rs(13).name & " name=" & rs(13).name & " value='" & rs(13) & "' class=dtn type='text' size=3 onblur='document.edGeneral.submit();' >"
		Else
			sInput = ofv.GenerarInput("",rs(13),"readonly","","1","dt","")
		End If
		Celda = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")



		'Descripcion
		Celda = ofv.GenCelda("","tt","","20%","","","","Descripcion","si")
		If  bModificar Then
			sInput = ofv.GenerarInput(rs(3).name,rs(3),"text","edGeneral","200","dt","")
			sInput = Replace(sInput,"{ document.edGeneral.submit(); return true;}","{ if (this.value.length > 70) {alert('El texto es demasiado extenso'); return false;} else { document.edGeneral.submit(); return true;}}")
		Else
			sInput = ofv.GenerarInput("",rs(3),"readonly","","90","dt","")
		End If
		Celda = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


		'Identificacion
		Celda = ofv.GenCelda("","tt","","","","","","Identificacion","si")
		If  bModificar Then
			sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","edGeneral","20","dt","")
		Else
			sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
		End If
		Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


		'PRODUCTO
		Celda = ofv.GenCelda("","tt","","","","","","PRODUCTO","si")

		If  bModificar Then
			Strsql = "Select * From TIPOPRODUCTO Order By descri"

			''  scomKK = "0" & chr(9)
			''  scomDD = "SIN DATOS" & chr(9)
			''''response.write strsql & "<br>" & chr(13)
			Set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do Until rsniv.eof
				scomKK = scomKK & rsniv(0) & Chr(9)
				scomDD = scomDD & rsniv(1) & Chr(9)
				RSNIV.MOVENEXT
			Loop
			Call ofv.cerrarconsulta(rsniv)

			sInput = ofv.gencomboFX("edGeneral",rs(9).name,rs(9).value,scomKK,scomDD,"sb0")
			''''''''sInput = ofv.GenerarCombo("edGeneral",rs(9).name,rs(9).value,Strsql,cn,0,1,"")
		Else
			If  CStr(rs(9) & "") <> "" Then
				Strsql = "Select descri From TIPOPRODUCTO Where objid = " & rs(9)
				'''''response.write strsql & "<br>" & chr(13)
				Set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If  Not rsInstrumento.eof Then sInstrumento = rsInstrumento(0)
				Call ofv.cerrarconsulta(rsInstrumento)
			End If
			sInput = ofv.GenerarInput("",sInstrumento,"readonly","","90","dt","")
		End If

		Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

		'Nivel de Confidencialidad
		Celda = ofv.GenCelda("","tt","","","","","","Nivel","si")

		If  bModificar Then

			scomKK = ""
			scomDD = ""
			'''''response.write strsql & "<br>" & chr(13)
			Set rsniv = ofv.crearconsultaEx(StrSqlx,cn2,1,parametros)
			Do Until rsniv.eof
				scomKK = scomKK & rsniv(0) & Chr(9)
				scomDD = scomDD & rsniv(1) & Chr(9)
				RSNIV.MOVENEXT
			Loop
			Call ofv.cerrarconsulta(rsniv)

			sInput = ofv.gencomboFX("edGeneral",rs(6).name,rs(6).value,scomKK,scomDD,"sb0")
			''  sInput = ofv.GenerarCombo("edGeneral",rs(6).name,rs(6).value,Strsqlx,cn2,0,1,"sb0")
		Else
			If  CStr(rs(6)) <> "" Then
				sNiv = rs(6)
				For ii=0 To UBound(CNIV)
					If  CStr(CNIV(ii)) = CStr(rs(6)) Then
						sNiv = DNIV(ii)
						Exit For
					End If
				Next
			End If
			sInput = ofv.GenerarInput("",sNiv,"readonly","","90","dt","")
		End If

		Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


		sInstrumento = ""
		'Reemplazo
		Celda = ofv.GenCelda("","tt","","","","","","Reemplaza a","si")
		If  rs(4) <> "" Then
			Strsql = "Select xdescri From csGenObjetos where codobjeto = " & rs(4)
			' response.write strsql & "<br>"
			Set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  Not rsInstrumento.eof Then sInstrumento = rsInstrumento(0)
			nDRCoboj = rs(0)
			nDRTiboj = rs(2)
			nObjrem  = rs(4)
			nXDescri = xdescri
			Call ofv.cerrarconsulta(rsInstrumento)
		End If
		sInput = sInstrumento

		Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

		If Nestado = 1 Then

			'DESTINO
			Celda = ofv.GenCelda("","tt","","","","","","Destinos","si")
			sImagenP = "&#0232;"
			sinput = "<a class='ft1n' href='edgeneraldest.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
			sinput = sinput  & "' title='Agregar destinos' >" & sImagenP & "</a>"

			Celda  = Celda & ofv.GenCelda("","uti","","","","","",sinput,"si")
			Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
		End If

		celda= ""

		strsql1 = "select * from DEPENREEMP where codobj = " & nCodigo
		Set rsOrigen1 = ofv.crearconsultaEx(strsql1,cn,1,parametros)

		Do While Not rsOrigen1.eof
			Strsql = "Select * From csObjetosRdo Where codobjeto = " & rsOrigen1(3)
			Set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			If Not rsOrigen.eof Then
				Celda = ofv.GenCelda("","","","","","","","","si")
				sInput = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6) & " - " & "(" & rsOrigen(7) & ")"
				Celda = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
				Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
			End If
			Call ofv.cerrarconsulta(rsOrigen)
			rsOrigen1.movenext
		Loop

		If  wkaplic = "PL" Then

			'Tipo de vinculacion

			SESSION("CODVINOBJ") = RS(17) & ""

			Celda = ofv.GenCelda("","tt","","","","","","Vinculacion","si")

			scomK = "S" & Chr(9) & "N" & Chr(9) & "X"
			scomD = "ESTRICTA" & Chr(9) & "REFERENCIAL" & Chr(9) & "SIN VINCULACION"

			WKVINC   = RS(17) & ""
			WKVINCOD = RS(15) & ""

			If WKVINCOD = "" Or WKVINCOD = "0" Then
				WKVINCOD = "NO"
			End If

			If WKVINC = "" Then WKVINC = "X"

			If bModificar And WKVINCOD <> "NO" And SESSION("CODVINOBJ") <> "SC" And SESSION("CODVINOBJ") <> "NC"  And SESSION("CODVINOBJ") <> "XC" Then

				sInput = ofv.gencomboFX("edGeneral",rs(17).name,WKVINC,scomK,scomD," ")
			Else
				If WKVINC = "S" Then
					WKVINC = "ESTRICTA"
				ElseIf WKVINC = "SC" Then
					WKVINC = "ESTRICTA"
				ElseIf WKVINC = "NC" Then
					WKVINC = "REFERENCIAL"
				ElseIf WKVINC = "N" Then
					WKVINC = "REFERENCIAL"
				Else
					WKVINC = "SIN VINCULACION"
				End If
				sInput = ofv.GenerarInput("",WKVINC,"readonly","","30","dt","")
			End If

			Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
			Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
		End If




		xestad  = 1
		Strsql  = "SELECT estado FROM OBJPRTY "
		Strsql  = Strsql & "where  codobjeto=" & nCodigo
		''response.write strsql
		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If  Not rsx.eof Then xestad = rsx(0)
		Call ofv.cerrarconsulta(rsx)

		sAccion = ""
		If bModificar Then
			sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objetos&volver=" & vuelta
			sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
		End If

		sTabla = ofv.GenTabla("","","","90%","","0","","2","0","","",Fila,"si")
		sHTML  = sHTML & ofv.GenForm("edGeneral","",sAccion,"","post","",sTabla,"si") & "&nbsp;"

		'NRO SOLICITUD
		If session("BASE") = "D" and session("APLMDCLI")="BGAL" Then

			If Nestado = 1 Then

				Celdax = ofv.GenCelda("","tt","","20%","","","","Nro. solicitud","si")
				sxInput = ofv.GenerarInput(rs(21).name,rs(21),"text","NroSolic","5","dt","")
				Celdax  = Celdax & ofv.GenCelda("","uti","","80%","","","",sxInput,"si")
				Filax   = ofv.GenRow("","","","","25","","",Celdax,"si")
				xTablax = ofv.GenTabla("","","","90%","","0","","2","0","","",Filax,"si")

			Else

				Celdax = ofv.GenCelda("","tt","","20%","","","","Nro. solicitud","si")
				sxInput = rs(21)
				Celdax  = Celdax & ofv.GenCelda("","uti","","80%","","","",sxInput,"si")
				Filax   = ofv.GenRow("","","","","25","","",Celdax,"si")
				xTablax = ofv.GenTabla("","","","90%","","0","","2","0","","",Filax,"si")

			End If

		End If

		If bModificar Then
			sAccion = "asp/ActualizarEN.asp?ID=" & rs("IDZ") & "&Tabla=objprty&volver=" & vuelta
			sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
		End If

		sHTML = sHTML & ofv.GenForm("NroSolic","",sAccion,"","post","",xTablax,"si")

		sHTML = sHTML & "<br><br><br>&nbsp;"

		Call ofv.cerrarconsulta(rs)

		'opciones
		'response.write bModificar
		If bModificar Then

			Fila =  ""

			celdas = ofv.GenCelda("","ttt colspan=5","","","","","","OPCIONES","si")
			filas  = ofv.GenRow("","","","","25","","",celdas,"si")


			' Alcance Empresas

			If  Session("EMPMUN") = "S"  And EsCopr = "S" Then

				sInput  = ofv.GenerarInput("bt1","Alcance","submit","","10","btsb","")
				'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
				sAccion = "EdgeneralEMPALC.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo
				bt1     = ofv.GenForm("bt1f","",sAccion,"","post","",sInput,"si")
				celdas1  = ofv.GenCelda("","ut","","","","","",bt1,"si")
				''    bt1     = ofv.GenForm("bt1f","",sAccion,"","post","",celdas,"si")
			Else
				celdas1     = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
				'bt1     = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
			End If

			' Reemplazo de documentos

			'response.write  nDRCoboj & " nDRCoboj " & "<br>"
			'response.write  nDRTiboj & " nDRTiboj " & "<br>"
			'response.write  nCodigo  & " nCodigo  " & "<br>"
			'response.write  sTipo    & " sTipo    " & "<br>"
			'response.write  nXDescri & " nXDescri " & "<br>"
			'response.write  nObjrem  & " nObjrem  " & "<br>"

			If CInt(xestad) = 1 Then
				sInput  = ofv.GenerarInput("bt2","Reemplazo","submit","","10","btsb","")
				sAccion = "EdgeneralBSSS.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SLetra=A&Sopcion=A&VASP=CM" & "&Objrem=" & nObjrem & "&nXDescri=" & nXDescri
				bt2     = ofv.GenForm("bt2f","",sAccion,"","post","",sInput,"si")
				celdas2  = ofv.GenCelda("","ut","","","","","",bt2,"si")
				'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
				'bt2     = ofv.GenForm("bt2f","",sAccion,"","post","",celdas,"si")
			Else
				celdas2  = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
				'bt2     = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
			End If
			' Agrupamiento del esquema
			sInput  = ofv.GenerarInput("bt3","Agrupamiento","submit","","10","btsb","")
			sAccion = "edgeneraladv.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo
			bt3     = ofv.GenForm("bt3f","",sAccion,"","post","",sInput,"si")
			celdas3  = ofv.GenCelda("","ut","","","","","",bt3,"si")
			'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
			'bt3     = ofv.GenForm("bt3f","",sAccion,"","post","",celdas,"si")

			' Palabras de busquedas
			sInput  = ofv.GenerarInput("boton","Palabras Busqueda","submit","","10","btsb","")
			sAccion = "EdgeneralPalBus.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
			bt4     = ofv.GenForm("bt4f","",sAccion,"","post","",sInput,"si")
			celdas4  = ofv.GenCelda("","ut","","","","","",bt4,"si")
			'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
			''    bt4     = ofv.GenForm("bt4f","",sAccion,"","post","",celdas,"si")

			bt5     = ""

			If  wkaplic = "PL" Then

				' Vinculacion
				sInput  = ofv.GenerarInput("boton","Vinculacion Normativa","submit","","10","btsb","")

				sAccion = "EdGENPL.asp?"
				sAccion = sAccion & "ID=" & nCodigo & "&TABLA=OBJETOS&Codobjeto=" & nCodigo
				sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vuelta2 & "_Codobjeto=" & nCodigo
				sAccion = sAccion & "_Tipobjeto=" & sTipo

				bt5     = ofv.GenForm("bt4f","",sAccion,"","post","",sInput,"si")
				celdas5  = ofv.GenCelda("","ut","","","","","",bt5,"si")
				'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
				'bt5     = ofv.GenForm("bt4f","",sAccion,"","post","",celdas,"si")

			End If


			' cambios al documento
			sInput  = ofv.GenerarInput("boton","Cambios al Documento","submit","","10","btsb","")
			sAccion = "EdgeneralCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
			bt6      = ofv.GenForm("bt5f","",sAccion,"","post","",sInput,"si")
			celdas6  = ofv.GenCelda("","ut","","","","","",bt6,"si")
			'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
			'bt6      = ofv.GenForm("bt5f","",sAccion,"","post","",celdas,"si")



			''' bt6     = ofv.GenCelda("","ut","","80%","","","","&nbsp;","si")


			' BLOQUE INFO
			sInput  = ofv.GenerarInput("boton","Informacion del documento","submit","","10","btsb","")
			sAccion = "Edgeneralinfo.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&family=D"
			bt7     = ofv.GenForm("bt7f","",sAccion,"","post","",sInput,"si")
			celdas7  = ofv.GenCelda("","ut","","","","","",bt7,"si")
			'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
			'bt7     = ofv.GenForm("bt7f","",sAccion,"","post","",celdas,"si")

			If session("APLMDCLI")="BP" Or session("APLMDCLI")="BGAL" Or session("APLMDCLI")="BC" Then
				sInput  = ofv.GenerarInput("boton","Eliminar Relaciones","submit","","10","btsb","")
				sAccion = "Eliminarlinks.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
				bt8     = ofv.GenForm("bt8f","",sAccion,"","post","",sInput,"si")
				celdas8  = ofv.GenCelda("","ut","","","","","",bt8,"si")
				'celdas  = ofv.GenCelda("","ut","","","","","",sInput,"si")
				'bt8     = ofv.GenForm("bt8f","",sAccion,"","post","",celdas,"si")
			Else
				BT8=""
			End If



			filas   = filas & ofv.GenRow("","","","","30","","",celdas1 & celdas2 & celdas3 & celdas4 & celdas5 & celdas6 & celdas7 & celdas8,"si")

			sHTML   = sHTML & ofv.GenTabla("","","","90%","","0","","2","0","","",filas,"si")

		End If


		'sHTML   = sHTML


	End If

	Agregar = ""
	Buscar  = ""
	Celda   = ""

	'cambios
	Listar = ""


	If bVersion = "SI" Then
		If SESSION("SHWMOD") <> "N" Then

		ElseIf SESSION("VIEWMOD") <> "N" Then

			'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
			Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

			sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
			sHTML = sHTML & ConsCambios(sasp)

			Celda = Listar


		End If


	End If







	Call ofv.cerrarconn(cn2)
	Call ofv.cerrarconn(cn)


	sHTML = sHTML & GTCOBS()


	'''''''
	sHTML = sHTML & "</span></div>"



	Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")

	sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")

	sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
	'  dim i
	'  For Each i in Session.Contents
	'    Response.Write i &"=" & session(i) & "<br>"
	'  Next



	sHTML = sHTML & "</body></html>"
End If







response.write sHTML

Function Fechador(wkfecha)



	fec = CDate(wkfecha + CLng(CDate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))

	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function


Function BotonEliminarlnk(sNombre, sHref,CLAVE)

	wktitdelete="Confirma la eliminacion de relaciones"
	sImageELS    =ofv.GenerarInput("boton","Eliminar links","button","","10","btsb","") '"<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
	sformobj   = "FRMOBJ" & CLAVE
	sAccion    = sHref
	formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
	sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "

	BotonEliminarlnk=ofv.GenCelda("","ut title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageELs ,"si")



End Function


%>
