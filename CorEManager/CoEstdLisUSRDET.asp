<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	
	
	
	Dim bconsultar()
	
	
	
	strsql = ""
	strsql1 = ""
	
	NCODIGO2 = REQUEST.querystring("USR")
	NCODIGO = REQUEST.querystring("CODIGO")
	STIPO   = REQUEST.querystring("TIPO")
	SCOD    = REQUEST.querystring("COD")
	
	
	XFDES = REQUEST.querystring("FDES")
	XFHAS = REQUEST.querystring("FHAS")
	
	XWHERE = " AND (FINICIO > " & (XFDES - 1) & " AND FINICIO < " & (XFHAS + 1) & ") AND USUARIO = '" & NCODIGO2 & "' "
	
	If SCOD = "SI" Then
		XWHERE = XWHERE & " AND CODOBJETO = " & NCODIGO & " AND TIPOBJETO = '" & STIPO & "' "
	Else
		XWHERE = XWHERE & " AND CODOBJETO > 0 "
	End If
	XORDER = " ORDER BY USUARIO, CODOBJETO, FINICIO DESC "
	
	
	
	strsql = strsql & "SELECT CODOBJETO, TIPOBJETO, FINICIO, HINICIO, HFINAL , "
	strsql = strsql & "(HFINAL-HINICIO) AS TIEMPO,"
	strsql = strsql & " USUARIO, USUARIO AS USERACC "
	strsql = strsql & " FROM OBJALOGDET B1"
	strsql = strsql & " WHERE TACC = 'CONSEN'  " & XWHERE
	
	strsql = strsql & XORDER
	
	'''''RESPONSE.WRITE STRSQL & "<BR>"
	
	
	
	' revisar porque se saco clng(cdate("01-01-1900") 4/7/2018 ale
	
	'sID  = XFDES  - clng(cdate("01-01-1900"))
	sID  = request.querystring("FDES")
	'sID2 = XFHAS  - clng(cdate("01-01-1900"))
	sID2 = request.querystring("FHAS")
	
	sTitulo = ""
	
	
	Set cn = OFV.conectar(OFV.strconn2)
	
	
	
	fi = CLng(CDate("01-01-1900"))
	
	
	
	sHTML = OFV.FormHeader(sTitulo)
	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
	xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)
	
	sHTML = sHTML & xs
	
	
	'   Celda = OFV.GenCelda("", "utg colspan=2 style={border-style:RIDGE;border-width:2;border-color:#F7AA22;} ", "", "", "", "", "", sTitulo, "si")
	
	'    Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	
	
	xxxx    = getInfoUserEX(NCODIGO2,session("un"))
	sGrpUsr = session("SGU_nombre")
	
	
	Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "USUARIO:", "si")
	Celda = Celda & OFV.GenCelda("", "uti", "", "80%", "", "", "", "" & sGrpUsr & "", "si")
	Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	
	
	'   Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "SECTOR:", "si")
	'   Celda = Celda & OFV.GenCelda("", "UTi", "", "80%", "", "", "", RSX(1), "si")
	'   Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	
	
	Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "DESDE:", "si")
	Celda = Celda & OFV.GenCelda("", "UTi", "", "80%", "", "", "", UCase(Fechador(sID)), "si")
	Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "HASTA:", "si")
	Celda = Celda & OFV.GenCelda("", "UTi", "", "80%", "", "", "", UCase(Fechador(sID2)), "si")
	Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	'' sTitulo = "ACCESO DE USUARIOS<BR> Desde: " & Fechador(sID) & "<BR> Hasta: " & Fechador(sID2)
	
	
	
	
	
	XTABLA = OFV.GenTabla("", "aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'", "", "95%", "", "0", "", "0", "0", "", "", Fila, "si")
	
	Fila = ""
	
	Celda = OFV.GenCelda("", "vt colspan=5", "", "", "", "", "", " ", "si")
	
	Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	
	Celda = ""
	Celda = Celda & OFV.GenCelda("", "TTr", "","50%", "", "", "", "DOCUMENTO", "si")
	Celda = Celda & OFV.GenCelda("", "TTr", "","25%", "", "", "", "FECHA", "si")
	Celda = Celda & OFV.GenCelda("", "TTr", "","25%", "", "", "", "ACCESOS", "si")
	
	Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	XTABLA2 = OFV.GenTabla("", "aaa style='border-radius:0px;border-bottom:none;'", "", "95%", "", "0", "", "0", "0", "", "", Fila, "si")
	
	sHTML = sHTML & "<CENTER>"
	sHTML = sHTML & "<div id='mt1' class='' style='display:block;width:650;' >"
	sHTML = sHTML & XTABLA & XTABLA2
	sHTML = sHTML & "</div>"
	
	Call OFV.cerrarconsulta(rsX)
	
	sHTML = sHTML & "<div id='mt2' class='' style='display:block;height:300px;width:650;overflow:auto;' >"
	
	
	Fila = ""
	
	
	
	
	Celda = OFV.GenCelda("", "UT colspan=5", "", "", "", "", "", " ", "si")
	
	Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	
	
	nLinea = 0
	
	
	PosReg = 0
	Postmp = 0
	Poscnt = 0
	
	PosReg2 = 0
	Postmp2 = 0
	
	
	SUSRANT = "NO"
	
	
	
	'response.write "sql: " & strsql & "<br>"
	Set rsX = OFV.crearconsultaEx(strsql, cn, 1, "")
	Do Until rsX.EOF
		'    RESPONSE.WRITE "U: " & RSX(6) & RSX(2) & "<BR>"
		
		sImagen  = "<i class='las la-file-alt' style=color:"& session("CMFPNTIT") &"></i>"
		sImagen2 = "<i class='las la-angle-right' style=color:"& session("CMFPNTIT") &"></i>"
		
		If SUSRANT <> CStr(RSX(0) & "") Then
			
			xTrColor = "#F9F9F9"
			
			If SUSRANT <> "NO" Then
				
				Celda = OFV.GenCelda("", "Ut COLSPAN=2 ", "RIGHT", "", "", "", "","&nbsp;", "si")
				Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "TOTAL&nbsp;", "si")
				Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "ACCESOS:&nbsp;" & PosReg2 & "&nbsp;", "si")
				Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "PROMEDIO:&nbsp;" & GETGLBHMS(Round((Postmp2/PosReg2),0)) & "&nbsp;", "si")
				
				Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")
				
			End If
			
			PosReg2 = 0
			Postmp2 = 0
			Poscnt  = Poscnt + 1
			
			SUSRANT = CStr(RSX(0))
			
			sTipobjeto = RSX(0)
			
			'   Strsql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & sTipobjeto
			'   set rsM = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
			'   IF NOT RSM.EOF THEN
			'     sTipobjeto1 = RSM(1) & ":&nbsp;"
			'     sTipobjeto2 = RSM(3) & " - "
			'     sTipobjeto3 = RSM(2) & " ("
			''     sTipobjeto4 = RSM(5) & " "
			'     sTipobjeto5 = RSM(6) & ") "
			'     sTipobjeto  = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
			'   END IF
			' sTipobjeto= EsqnomP("", sTipobjeto)
			sTipobjeto= EsqnomPx(RSX("tipobjeto"), RSX("codobjeto"),ofv)
			Celda = OFV.GenCelda("", "UTS", "", "2%", "", "", "", sImagen, "si")
			Celda = Celda & OFV.GenCelda("", "VT COLSPAN=4 ", "", "", "", "", "", "&nbsp;" & sTipobjeto, "si")
			
			Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")
		End If
		
		
		dVigencia = "SD"
		If Not IsNull(rsX(2)) Then dVigencia = CDate(CLng(rsX(2)))
		dVign = InStr(dVigencia, "1900")
		If dVign > 0 Then dVigencia = "SD"
		
		
		XTIME =  RSX(5) & ""
		
		If XTIME = "" Or XTIME = "0" Then XTIME = "1"
		
		XTIME =  CLng(XTIME)
		
		
		Celda = OFV.GenCelda("", "AAA COLSPAN=2 ", "RIGHT", "", "", "", "","&nbsp;", "si")
		Celda = Celda & OFV.GenCelda("", "UTS", "RIGHT", "", "", "", "",sImagen2, "si")
		Celda = Celda & OFV.GenCelda("", "UT  bgcolor='" & xTrColor & "' ", "RIGHT", "", "", "", "", FechadorCor(dVigencia,0) & "&nbsp;", "si")
		Celda = Celda & OFV.GenCelda("", "UT  bgcolor='" & xTrColor & "' ", "RIGHT", "", "", "", "", GETGLBHMS(XTIME) & "&nbsp;", "si")
		
		Fila = Fila & OFV.GenRow("", "AAA", "", "", "", "", "", Celda, "si")
		
		If  xTrColor = "#F9F9F9" Then
			xTrColor = "#F3F3F3"
		Else
			xTrColor = "#F9F9F9"
		End If
		
		
		
		PosReg = PosReg + 1
		'     PosReg2 = PosReg2 + RSX(8)
		PosReg2 = PosReg2 + 1
		
		Postmp = Postmp + XTIME
		
		Postmp2 = Postmp2 + XTIME
		
		rsX.movenext
		
		
	Loop
	
	
	
	
	If PosReg > 0 Then
		
		
		
		If SUSRANT <> "NO" Then
			
			Celda = OFV.GenCelda("", "vt COLSPAN=2 ", "RIGHT", "", "", "", "","&nbsp;", "si")
			Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "TOTAL&nbsp;", "si")
			Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "ACCESOS:&nbsp;" & PosReg2 & "&nbsp;", "si")
			Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "PROMEDIO:&nbsp;" & GETGLBHMS(Round((Postmp2/PosReg2),0)) & "&nbsp;", "si")
			
			Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")
			
		End If
		
		
		
		Celda = OFV.GenCelda("", "vt COLSPAN=2 ", "", "", "", "", "","TOTAL GENERAL&nbsp;", "si")
		Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "DOCUMENTOS:&nbsp;" & Poscnt, "si")
		Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "ACCESOS:&nbsp;" & PosReg & "&nbsp;", "si")
		Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "PROMEDIO:&nbsp;" & GETGLBHMS(Round((Postmp/PosReg),0)) & "&nbsp;", "si")
		
		FilaF = OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")
		XTABLAF = OFV.GenTabla("", "aaa style='border-radius: 0px 0px 10px 10px;'", "", "95%", "", "0", "", "-1", "0", "", "", FilaF, "si")
		
		
		
	Else
		
		Celda = ofv.GenCelda("","vt  colspan=5 ","","","","",""," NO HAY ACCESOS PARA LA PRESENTE SELECCION ","si")
		Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
		
		
	End If
	
	Celda = OFV.GenCelda("", "UT", "", "2%", "", "", "","&nbsp;", "si")
	Celda = Celda & OFV.GenCelda("", "UT", "", "43%", "", "", "", "&nbsp;", "si")
	Celda = Celda & OFV.GenCelda("", "UT", "", "10%", "", "", "", "&nbsp;", "si")
	Celda = Celda & OFV.GenCelda("", "UT", "", "20%", "", "", "", "&nbsp;", "si")
	Celda = Celda & OFV.GenCelda("", "UT", "", "25%", "", "", "", "&nbsp;", "si")
	Fila  = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	
	XTABLA = OFV.GenTabla("", "aaa style='border-radius: 0px;border-bottom:none;'", "", "95%", "", "0", "", "-1", "0", "", "", Fila, "si")
	sHTML  = sHTML & XTABLA
	sHTML  = sHTML &  "</div>"
	
	sHTML = sHTML & "<div id='mtF' class='' style='display:block;width:650;' >"
	sHTML  = sHTML & XTABLAf
	sHTML  = sHTML &  "</div>"
	sHTML  = sHTML & "</CENTER>"
	sHTML  = sHTML &  "</body></html>"
	
	Call OFV.cerrarconsulta(rsX)
	Call OFV.cerrarconn(cn)
	
	
	
End If


response.write sHTML
''response.flush

Function EsqnomP(xtipo, xcodigo)
	
	
	
	
	If xtipo = "RSK" Then
		
		StrSql = "SELECT * FROM  ORRSKCIRCUITO "
		StrSql = StrSql & " Where OBJID = 0" & xcodigo
		Set xrs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
		If Not xrs.EOF Then
			If xrs("APLIC") = "OR" Then
				xdescri = "CIRCUITO AUTOEVALUACION"
			Else
				xdescri = "PLAN DE CONTINGENCIAS"
			End If
			
			StrSql = "Select * From objetos where codobjeto = 0" & xrs("CODOBJETO")
			Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
			If Not xrs.EOF Then
				xiden = xrs(1)
				xobjdes = xrs(3)
			End If
			
			
		End If
		Call ofv.cerrarconsulta(xrs)
		
		sTipobjetoX = Capitalize(xdescri) & ": " & Capitalize(xobjdes)
		
	Else
		
		
		sTipobjetoX = xcodigo
		
		StrSql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
		Set RSM = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)
		If Not RSM.EOF Then
			
			
			
			
			If session("APLMDCLI") ="BC" Then
				sTipobjeto1 = (RSM(1)) & ":&nbsp;"
				sTipobjeto2 = RSM(3) & ""
				sTipobjeto3 = " - " & (RSM(2) & "")
				sTipobjeto4 = " (" & (RSM(5) & "") & " "
				sTipobjeto5 = (RSM(6) & "") & ") "
			Else
				sTipobjeto1 = Capitalize(RSM(1)) & ":&nbsp;"
				sTipobjeto2 = Capitalize(RSM(3) & "")
				
				sTipobjeto3 = " - " & Capitalize(RSM(2) & "")
				sTipobjeto4 = " (" & Capitalize(RSM(5) & "")
				If UCase(RSM(5) & "")="O" Then
					sTipobjeto5 = ")"
				Else
					sTipobjeto5 =  " " & Capitalize(RSM(6) & "") & ") "
				End If
			End If
			sTipobjeto6 =  " (" & Capitalize(RSM(7) & "") & ") "
			If session("FULLDOCN") = "1" Then
				sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto6
			ElseIf session("FULLDOCN") = "2" Then
				sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto6
			Else
				sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5 & sTipobjeto6
			End If
			
			
			
		End If
		Call ofv.cerrarconsulta(RSM)
	End If
	
	
	
	
	
	
	
	
	EsqnomP = sTipobjetoX
	
	
End Function


Function Fechador(wkfecha)
	
	
	
	fec = CDate(wkfecha) ' + clng(cdate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))
	
	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)
	
	
End Function

%>