<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires = 0
	CantRegMostrarx =10
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	ofv.ObtenerAtributos Session("Form"),""
	'bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	'bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar
	APLCHECK()
	base=request.querystring("BASE")
	Filtro=request.querystring("filtro")
	If Filtro <> "" Then
		session("Tfiltro")=filtro
	Else
		filtro=session("Tfiltro")
	End If


	MasterType = "SCR"
	Mtype = request.querystring("MType") & ""

	If Mtype <> "" Then MasterType = Mtype & ""

	If MasterType <> "EXP" Then
		sHTML = OFV.FormHeader(Session("FormN")) & Chr(13)
	Else
		Response.ContentType = "application/vnd.ms-excel"

		Select Case filtro
			Case "NV"
			 Response.AddHeader "content-disposition", " filename=NoVencidos.xls"
			Case "V"
			 Response.AddHeader "content-disposition", " filename=Vencidos.xls"
			Case Else
			 Response.AddHeader "content-disposition", " filename=salida.xls"
		End Select
	End If

	If base <> "" Then
		SESSION("BASE")=base
	End If


   if SESSION("BASE")="K" then
      tdoc="Formularios"
   else
      tdoc="Documentos"
   end if

	stilde="<font face=corflaticon size=4>a</font>"

	If MasterType <> "EXP" Then

		If filtro="T" Then
			slink="<a class=ut href=VtoRevLis.asp?filtro=T target=_self ><B><u>Todos Los " & tdoc & "&nbsp;&nbsp;</u></B> " & stilde &"</A>"
		Else
			slink="<a class=ut href=VtoRevLis.asp?filtro=T target=_self ><B>Todos Los " & tdoc & "&nbsp;&nbsp;</B></A>"
		End If

		Celda =  ofv.GenCelda("","vt","left","","","","",slink,"si")
		If filtro="NV" Then
			slink="<a class=ut href=VtoRevLis.asp?filtro=NV target=_self ><B><u>" & tdoc & " No Vencidos&nbsp;&nbsp;</u></B>"& stilde & "</A>"
		Else
			slink="<a class=ut href=VtoRevLis.asp?filtro=NV target=_self ><B>" & tdoc & " No Vencidos&nbsp;&nbsp;</B></A>"
		End If

		Celda =  celda & ofv.GenCelda("","vt","left","","","","",slink,"si")
		If filtro="V" Then
			slink="<a class=ut href=VtoRevLis.asp?filtro=V target=_self ><B><u>" & tdoc & " Vencidos&nbsp;&nbsp;</u></B>" & stilde & "</A>"
		Else
			slink="<a class=ut href=VtoRevLis.asp?filtro=V target=_self ><B>" & tdoc & " Vencidos&nbsp;&nbsp;</B></A>"
		End If

		Celda =  celda & ofv.GenCelda("","vt","left","","","","",slink,"si")
		Fila  = ofv.GenRow("","","","","","","",celda,"si")
		tablax= ofv.GenTabla("","","center","50% ","","0","","0","1","","",fila,"si")
		shtml = shtml & tablax & "<br>"
	End If
	nFecini = 0

	If nFecini < CLng(CDate("01-01-1900")) Then
		nFecini = CLng(CDate("01-01-1900"))
	Else
		nFecini = 0
	End If

	xfnul = CLng(CDate("01-01-1900")) - nFecini
	xfhoy = CLng(Date()) - nFecini

	Set cn = ofv.conectar(ofv.strconn2)

	'********************************************************
	'************ Variables para los colores del semaforo
	'*******************************************************

	strVerde = "SELECT * FROM diasalerta "
	strVerde = strVerde & "WHERE color = 'Green'"

	Set rsVerde = ofv.crearconsultaEx(strVerde,cn,1,parametros)

	MayorVerde = rsVerde(2)
	MenorVerde = rsVerde(3)

    Call ofv.cerrarconsulta(rsVerde)
	'response.write "Mayor:" & MayorVerde & "Menor:" & MenorVerde

	strAmarillo = "SELECT * FROM diasalerta "
	strAmarillo = strAmarillo & "WHERE color = 'Yellow'"

	Set rsAmarillo = ofv.crearconsultaEx(strAmarillo,cn,1,parametros)

	MayorAmarillo = rsAmarillo(2)
	MenorAmarillo = rsAmarillo(3)

    Call ofv.cerrarconsulta(rsAmarillo)
	'response.write "Mayor:" & MayorAmarillo & "Menor:" & MenorAmarillo


	strRojo = "SELECT * FROM diasalerta "
	strRojo = strRojo & "WHERE color = 'Red'"

	Set rsRojo = ofv.crearconsultaEx(strRojo,cn,1,parametros)

	MayorRojo = rsRojo(2)
	MenorRojo = rsRojo(3)

    Call ofv.cerrarconsulta(rsRojo)

	Celda = OFV.GenCelda("", "tt", "center", "", "", "", "", "DESCRIPCION", "si")
	IF session("APLMDCLI")="BGAL" THEN
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "APROBADORES", "si")
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "<center>FECHA<br><center>PUBLICACION", "si")
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "<center>FECHA<br><center>REVISION", "si")

	ELSE
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "TIPO", "si")
	Celda = celda & OFV.GenCelda("", "tt","center", "", "", "", "", "ESTADO", "si")
	' Celda = celda & OFV.GenCelda("", "tt", "", "", "", "", "", "AUTOR", "si")
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "<center>FECHA<br><center>PUBLICACION", "si")
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "<center>FECHA<br><center>REVISION", "si")
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "<center>FECHA<br><center>ACTUAL", "si")
	END IF
	Celda = celda & OFV.GenCelda("", "tt", "center", "", "", "", "", "<center>DIFERENCIA<br><center>DIAS", "si")

	Fila =   OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

	strFVto = "SELECT b.descri, a.tipobjeto, c.descri, c.tipocpt, a.estado, d.estado, a.codobjeto, a.fecvto, a.fecpublic, (cast(a.fecvto as int) -" & (xfhoy) & ") as diferencia, a.version, a.nrover, a.autor, a.codobjeto, e.TEMADESCRI, e.stemadescri "
	'cast(a.fecpublic AS int)) as diferencia, a.version, a.nrover "
	strFVto = 	strFVto & "FROM Objprty a "
	strFVto = 	strFVto & " INNER JOIN OBJETOS b ON ( a.codobjeto = b.codobjeto and a.tipobjeto = b.tipobjeto) "

	If SESSION("BASE")="K" Then
		strFVto = 	strFVto & " INNER JOIN tipOBJETO c ON ( a.tipobjeto = c.tipobjeto and c.TIPOCPT = 'K') "
	Else
		strFVto =   strFVto & " INNER JOIN tipOBJETO c ON ( a.tipobjeto = c.tipobjeto and c.TIPOCPT = 'D') "
	End If
	strFVto = 	strFVto & " INNER JOIN Estados d ON ( a.estado = d.objid ) "
	strFVto = 	strFVto & " INNER JOIN csrelpathdescriex e ON ( a.codobjeto = e.cobj and  a.tipobjeto = e.tobj)"
	strFVto = 	strFVto & " WHERE a.fecvto IS NOT NULL "

	'strFVto = 	strFVto & " AND a.estado=3 and a.fecvto <> 0 AND ((cast(a.fecvto as int) - " & (xfhoy) & ") <101 and (cast(a.fecvto as int) -" & (xfhoy) & ") > -101)"

	strFVto = 	strFVto & " AND a.estado=3 and a.fecvto <> 0 "

	'RESPONSE.WRITE strFVto

	'strFVto = 	strFVto & " AND c.TIPOCPT = 'D' "

	Select Case filtro
		Case "NV"
		 strFVto = 	strFVto & " and (cast(a.fecvto as int) -" & (xfhoy) & ") > = 0 "
		Case "V"
		 strFVto = 	strFVto & " and (cast(a.fecvto as int) -" & (xfhoy) & ") <  0 "
		Case Else
	End Select

	strFVto = 	strFVto & " order by diferencia"

	'Response.Write strFVto

	Set rsFVto = ofv.crearconsultaEx(strFVto,cn,1,parametros)

	'*************'
	'*** Variable para definir tema subtema ***'
	'***********'

	'	if Not rsFVto.Eof then

		'response.write temasubtema

	'	end if


	'response.write rsFVto(0)
	If Not rsFVto.eof Then rsFVto.move(CantRegMoverx)

	posi = 0

	Do While (Not rsFVto.eof) And (posi < CantRegMostrarx)
        TemaSubtema = "TEMA: " & rsFVto(14) & " / SUBTEMA: " & rsFVto(15)


		Celda = OFV.GenCelda("", "UT","", "", "", "", "", "<span title='" & TemaSubtema & "'>" & EsqnomPx(rsFVto("tipobjeto"), rsFVto("codobjeto"), ofv) & "</span>", "si")

		IF session("APLMDCLI")="BGAL" THEN
		 sqlref="select usr from OBJWFLDET where codobjeto='" & rsFVto("codobjeto") & "' and tipobjeto='" & rsFVto("tipobjeto") & "'"
         Set rsREF = ofv.crearconsultaEx(sqlref,cn,1,parametros)
         DELIM=""
         StrAprob=""
         DO WHILE NOT rsREF.EOF
          xx=getinfouser(rsREF("usr"))
          StrAprob=StrAprob & delim & session("SGU_nombre")
          delim="/"
          rsREF.movenext
         loop
         call ofv.cerrarconsulta(rsREF)
		 Celda = celda & OFV.GenCelda("", "UT", "", "", "", "", "", StrAprob, "si")
		 fpublic = CDate(rsFVto(8) + nFecini)
	     Celda = celda & OFV.GenCelda("", "UT", "center", "", "", "", "",  fechadorCor(fpublic, ""), "si")
		 frev = CDate(rsFVto(7) + nFecini)
	     Celda = celda & OFV.GenCelda("", "UT", "center", "", "", "", "", fechadorCor(frev, ""), "si")

		ELSE
		 Celda = celda & OFV.GenCelda("", "UT", "", "", "", "", "", rsFVto(2), "si")
		 Celda = celda & OFV.GenCelda("", "UT", "center", "", "", "", "", rsFVto(5), "si")
		' Celda = celda & OFV.GenCelda("", "UT", "", "", "", "", "", rsFVto(12), "si")
		 fpublic = CDate(rsFVto(8) + nFecini)
		'fechadorL(CDate(rsFVto(8) + nFecini))& " ---" &
		 Celda = celda & OFV.GenCelda("", "UT", "center", "", "", "", "",  fechadorCor(fpublic, ""), "si")
		 frev = CDate(rsFVto(7) + nFecini)
		 Celda = celda & OFV.GenCelda("", "UT", "center", "", "", "", "", fechadorCor(frev, ""), "si")
	     fhoy = Date()
	   	 Celda = celda & OFV.GenCelda("", "UT", "center", "", "", "", "",  fechadorCor(fhoy, ""), "si")

	   END IF


		'**************************************
		'**** Variables para definir Color ****
		'**************************************


		If  (rsFVto(9) > MayorVerde And rsFVto(9) <= MenorVerde ) Then
		Celda = celda & OFV.GenCelda("", "tt style='vertical-align:middle;font-family:verdana;font-size:8pt;font-weight:bold;background-color:green;border-radius:10px;border-style:solid;border-color:white;border-width:1;color:white'", "center", "", "", "", "", rsFVto(9), "si")
		End If

		If (rsFVto(9) > MayorAmarillo And rsFVto(9) <= MenorAmarillo ) Then
		  Celda = celda & OFV.GenCelda("", "tt style='vertical-align:middle;font-family:verdana;font-size:8pt;font-weight:bold;background-color:yellow;border-radius:10px;border-style:solid;border-color:white;border-width:1;color:black;'", "center", "", "", "", "", rsFVto(9), "si")
		End If

		If (rsFVto(9) > MayorRojo And rsFVto(9) <= MenorRojo ) Then
			Celda = celda & OFV.GenCelda("", "tt style='vertical-align:middle;font-family:verdana;font-size:8pt;font-weight:bold;background-color:red;border-radius:10px;border-style:solid;border-color:white;border-width:1;color:white'", "center", "", "", "", "", rsFVto(9), "si")
		End If

		If xTrColor = "#FFFFFF" Then
		   xTrColor = "#F0F0F0"
	    Else
		   xTrColor = "#FFFFFF"
        End If

		Fila =   FILA & OFV.GenRow("aaa bgcolor=" & xTrColor & " ", "", "", "", "25", "", "", Celda, "si")
		If MasterType <> "EXP" Then
			posi = posi + 1
		End If



		rsFVto.movenext
	Loop
	xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Exportar", "NO", "Exportar")
	sAccion = "vtoRevlis.asp?Mtype=EXP&filtro=" & filtro
	Listar  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")
	sHTML = sHTML  & ofv.GenTabla("","aaa","","90%","","0","0","0","0","","",FILA,"si")

	If MasterType <> "EXP" Then
		sHTML = sHTML & ofv.botoneraESSX("2",rsFVto,cantregmoverx,cantregmostrarx,"no","no",posi,"","","","",listar,"")
	End If


	Call ofv.cerrarconsulta(rsFVto)
	Call ofv.cerrarconn(cn)
	'response.write "</tr>"
	response.write SHTML

End If



Function FechadorL(Fecha)

	' On Error GoTo FechadorXErr

	Dim ndia,  nMes,  nAno

	'ndia = Mid(Fecha, 4, 2)
	'nMes = Left(Fecha, 2)
	ndia = Left(Fecha, 2)
	nMes = Mid(Fecha, 4, 2)
	nAno = Right(Fecha, 4)

	FechadorL = ndia & "-" & nMes & "-" & nAno

End Function

Function fechaConvert(wkfecha)
	Nfecha=CDate(wkfecha)
	Nfecha=DateSerial(Year(wkfecha),Month(wkfecha),Day(wkfecha))
	fechaConvert=Nfecha

End Function


%>
