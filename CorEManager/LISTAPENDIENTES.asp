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
	vuelta 		= 	ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

	SPROCESO 	= 	REQUEST.QUERYSTRING("PROCESO")
	SFILTRO  	= 	REQUEST.QUERYSTRING("FILTRO")
	SFILSEL  	= 	REQUEST.QUERYSTRING("FILSEL")
	SsubFILTRO  = 	REQUEST.QUERYSTRING("SUBFILTRO")
	SsubFILTRO2 = 	REQUEST.QUERYSTRING("SUBFILTRO2")
	SsubFILTRO3 = 	REQUEST.QUERYSTRING("SUBFILTRO3")
	XAPL		=	REQUEST.QUERYSTRING("APL")

	sMenu = request.querystring("Menu")
	sNombre = request.querystring("Nombre")
	If  XAPL<> "" Then
		session("XAPL")=XAPL
	Else
		XAPL=session("XAPL")
	End If
	'response.write xapl

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
	 xs = "<script language=JavaScript>" & Chr(13)
	 xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	 xs = xs & " var h=570;w=790 ;"

	If  session("APLMDCLI")="BC" Then
		xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If

	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & ","' & Chr(34)
	'xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
	xs = xs &"  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'" '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

	sHTML = sHTML & xs


	ofv.ObtenerAtributos Session("Form"),""
	bConsultar = ofv.AConsultar

	Set cn = ofv.conectar(ofv.strconn2)

	'STRSQL = "SELECT TIPOBJETO  FROM TIPOBJETO WHERE APLIC IN ('EM','GLB')"
'response.write XAPL

	If SPROCESO = "0" Then
		If XAPL="C" Then
			SQL= "SELECT * FROM TIPOBJETO  WHERE APLIC IN ('EM','GLB') AND TIPOCPT IN('D','C') "
		ElseIf XAPL="PD" Then
			SQL= "SELECT * FROM TIPOBJETO  WHERE APLIC IN ('EM','GLB') AND TIPOCPT IN('PD') "
		ElseIf XAPL="K" Then
			SQL= "SELECT * FROM TIPOBJETO  WHERE APLIC IN ('EM','GLB') AND TIPOCPT IN('K') "
		Else
			SQL= "SELECT * FROM TIPOBJETO "
		End If

		'response.write sql

		Set rs = ofv.crearconsultaEx(SQL,cn,1,parametros)

		DELIM=""
		STIPOS=""
		Do While Not RS.EOF

			STIPOS=STIPOS & DELIM & "'" & RS(0) & "'"
			DELIM=","

			RS.MOVENEXT
		Loop

		STIPOS= " IN(" & STIPOS & ")"
	Else
		STIPOS= " IN('" & SPROCESO & "')"
	End If



	Fila = ""
	xTrColor = "#F9F9F9"


	sImagevMa     = "<font face='arial' size=1 xcolor='#aa0000' style='font-weight:900;background-color:#aa0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MUY&nbsp;ALTO&nbsp;</font>"
	sImageva      = "<font face='arial' size=1 xcolor='#f7aa22' style='font-weight:900;background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ALTO&nbsp;</font>"
	sImagevm      = "<font face='arial' size=1 xcolor='#f7EEDD' style='font-weight:900;background-color:#f7EEDD;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MODERADO&nbsp;</font>"
	sImagevb      = "<font face='arial' size=1 xcolor='#6699CC' style='font-weight:900;background-color:#6699CC;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ESCASO&nbsp;</font>"
	sImagevNUL    = "<font face='arial' size=1 xcolor='#00AA00' style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NULO&nbsp;</font>"

	sImageSI      = "<font face='arial' size=1  style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;SI&nbsp;</font>"
	sImageNO      = "<font face='arial' size=1  style='font-weight:900;background-color:#AA0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NO&nbsp;</font>"

	sImagevn      = "<font face='arial' size=1 xcolor='#dddddd' style='font-weight:900;background-color:#dddddd;color:#ffffff' >----</font>"

	wkstylof      = "AAA style='font-family:arial;font-size:6pt;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' "
	wkstylof2     = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "

	''''b.codvinobj = 'VN' AND

	Set cnx = ofv.conectar(ofv.strconn5)

	Strsql = ""

      Strsql = Strsql & " SELECT a.*  FROM CKTVIEWex a "
     ' se agrego que no tome el paso 999
      Strsql = Strsql & " inner join wkfpasos b1 on(b1.wkfcod=a.model and b1.wkfpascod=a.pscmodel and b1.fin <> 1 ) "

	If session("EMPMUN") = "S" Then

		Strsql = Strsql & " inner join    (select distinct codempre,codigo,segckt from segcktresp  ) b "
		Strsql = Strsql & " on (b.codigo=a.pscresponsable and b.segckt=a.segckt) "

	End If

	Strsql = Strsql & " WHERE OBJTIPO " & STIPOS & " AND "
  'response.write sproceso
	If CStr(sproceso) <> "0" Then

		Strsql = Strsql & "  objtipo ='" & sproceso & "'  AND "

	End If

	If CStr(SsubFILTRO) <> "0" Then

		Strsql = Strsql & "  PSCRESPONSABLE ='" & SsubFILTRO & "'  AND "

	End If


      Strsql = Strsql & " a.FIN = 'NO' AND a.PSCFIN = 'NO' "

	If session("EMPMUN") = "S" Then

		Strsql = Strsql & " and a.codempre=" & session("UN")

	End If
    '  Strsql = Strsql & " ORDER BY a.SEGCKT DESC, PSCRESPONSABLE, PSCCOMIENZO "

      Strsql = Strsql & " ORDER  by fechadif,objcod,objtipo"

	'  response.write strsql & "<br>"
	Set rs = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
	If rs.eof Then

		Celda = ""
		Celda = Celda & ofv.gencelda("","VTG colspan=9 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY CIRCUITOS PENDIENTES PARA<BR>LA BUSQUEDA SOLICITADA","si")
		Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")

	Else
		posI = 0

		wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;' "


		WKIPROCc = 0
		MAINPROC = 0
		MAINACTI = ""

		Celda = ""
		If session("APLMDCLI") ="BGAL" Then
			Celda = Celda & ofv.gencelda("","ttt colspan=8 style='color:#ffffff' " & wkstylo1,"","","","","","LISTA DE CIRCUITOS PENDIENTES DEL WORKFLOW","si")
		Else
			Celda = Celda & ofv.gencelda("","ttt colspan=7 style='color:#ffffff' " & wkstylo1,"","","","","","LISTA DE CIRCUITOS PENDIENTES DEL WORKFLOW","si")
		End If

		Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

		Celda = ""
		Celda = Celda & ofv.gencelda("","ttt " & wkstylo1,"","","","","","DUEÑO","si")
		Celda = Celda & ofv.gencelda("","ttt " & wkstylo1,"","","","","","CIRCUITO","si")
		Celda = Celda & ofv.gencelda("","ttt " & wkstylo1,"","","","","","PASO","si")
		Celda = Celda & ofv.gencelda("","ttt " & wkstylo1,"","","","","","DOCUMENTO","si")

		If session("APLMDCLI") ="BGAL" Then
			Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","","","","","ACCESO WKF","si")
		End If

		Celda = Celda & ofv.gencelda("","ttt " & wkstylo1,"","","","","","DEMORA / DIAS","si")
		Celda = Celda & ofv.gencelda("","ttt " & wkstylo1,"","","","","","INICIO","si")
		Celda = Celda & ofv.gencelda("","ttt " & wkstylo1,"","","","","","INGRESAR&nbsp;PASO","si")
		Fila  = FILA & ofv.genrow("","aaa bgcolor='#ddddff'  ","","","25","","",Celda,"si")

		If Not rs.eof Then rs.move(CantRegMoverx)


		Do While (Not rs.EOF) And (POSI < CantRegMostrarx)
' response.write rs(10).name & "  222"
            if isdate(rs(10) & "") then

			 sInput2 = cdate(rs(10))
			else
			  sInput2 = cdate("1-1-1900")
			end if

			'sacamos la hora de la fecha ya que influye en la cantidad de dias AM/PM y suma 1 dia mas al convertir a clng
            sInput2= dateserial(year(sInput2),month(sInput2),day(sInput2))
			sinput3 = clng(date()) - clng(sInput2)


		'	sinputxx = CLng(Date())
		'	sinputzz = CLng(sInput2)
		'	sinputyy = Date()


			smarca = "S"


			'''#F7DD88

			'''#F7DD88
			Strsql = ""
			sql="select tipocpt from tipobjeto where tipobjeto='" & rs("objtipo") &"'"
			Set rsX1 = ofv.crearconsultaEx(sql,cn,1,parametros)
			Xtip="C"
			If Not  rsX1.eof Then Xtip=rsx1(0)
			Call ofv.cerrarconsulta(rsX1)
			sInputn=""
			If Xtip="PD" Or Xtip="SN" Then

				Strsql = Strsql & " SELECT 'Solicitud Nro: ' + cast (a.objid as nvarchar)  + '   ' + b.nombre "
				Strsql = Strsql & " FROM pedidos A "
				Strsql = Strsql & " inner join pedidosdet b on (a.objid=b.idpedido) "
				Strsql = Strsql & " where a.objid = 0" & RS(4)
				' response.write strsql
				sInputn=""
				Set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsX1.eof Then sInputn = RSX1(0)
				Call ofv.cerrarconsulta(rsX1)
			Else
				Strsql = Strsql & " SELECT A.CODOBJETO, (A.DESCRI + ': ' + A.OBJDESCRI) AS NOMBRE,tipobjeto "
				Strsql = Strsql & " FROM CSOBJETOSRDO A "
				Strsql = Strsql & " where a.codobjeto = 0" & RS(4) &  " and  a. tipobjeto='" &  rs("objtipo") & "'"
				Strsql = Strsql & " ORDER BY NOMBRE "
				'response.write strsql & chr(13) & "<br>"
				sInputn=""
				Set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsX1.eof Then sInputn = RSX1(1)
				Call ofv.cerrarconsulta(rsX1)
			End If

			Celda = ""
           if sInputn <> "" then
			sInput = RS(14)

			sql = "select distinct codempre from segcktresp where segckt='" & rs(5) & "' and codigo='" & rs(6) & "'"
			'response.write sql
			Set rx = ofv.crearconsultaEx(sql,cnX,1,parametros)

			codempre=0
			If Not rx.eof Then codempre = rx(0)

			xxxx = getInfoUserEX(RS(14),codempre)

			If session("SGU_nombre") <> "" Then    sInput =  session("SGU_nombre")


			Celda = Celda & ofv.gencelda("","UTi style='border-style:solid;border-width:1;border-color:#aaaaaa;color:" & Session("CMFPNTITFT") & ";background-color:color-mix(in srgb, " & Session("CMFPNTIT") & ", transparent 20%);'","","","","","",sInput,"si")

			sInput = RS(7)
			Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")

			sInput = RS(16)
			Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")


			sInput = RS(4)


			sHref = "&nbsp;"
			If RS("MODEL")="WKF00016631121722600921928" Then
				sActcon = "edobjwfldetPEND.asp?Codobjeto=" & RS("OBJCOD") & "&Tipobjeto=" & RS("OBJTIPO") & "&acc=PEND&ckt=AUTO&segckt=" & rs("segckt")
	 ELSEIF RS("MODEL")="WKF00016631121722600921928DD" THEN
      sActcon = "edobjwfldetPEND.asp?Codobjeto=" & RS("OBJCOD") & "&Tipobjeto=" & RS("OBJTIPO") & "&acc=PEND&ckt=AUTOSCAN&segckt=" & rs("segckt")
			ElseIf RS("MODEL")="WKF00016631121722600921928PED" Or RS("MODEL")="WKF00016631121722600921928PEDSOL" Then
				sActcon =""' "edobjwfldetPEND.asp?Codobjeto=" & RS("OBJCOD") & "&Tipobjeto=" & RS("OBJTIPO") & "&acc=PEND&ckt=PEDIDO&segckt=" & rs("segckt")
			Else
				sActcon = "edobjwfldetPEND.asp?Codobjeto=" & RS("OBJCOD") & "&Tipobjeto=" & RS("OBJTIPO") & "&acc=PEND&ckt=APRO&segckt=" & rs("segckt")
			End If

			'Strsql = Strsql & " SELECT A.CODOBJETO, (A.DESCRI + ': ' + A.OBJDESCRI) AS NOMBRE "
			'Strsql = Strsql & " FROM CSOBJETOSRDO A "
			'Strsql = Strsql & " where a.codobjeto = 0" & RS(4) & " and a.tipobjeto='" & RS("OBJTIPO") & "'"
			'Strsql = Strsql & " ORDER BY NOMBRE "
'
          '  response.write Strsql
			'Set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			'If Not RSX1.eof Then
			'sInput = RSX1(1) & ""
			'end if
'
			'Call ofv.cerrarconsulta(rsX1)

			'edobjwfldet.asp?volver=edobjwfl.asp&Codobjeto=2162&Tipobjeto=9

			sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)

			sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion,   sinputn, "si")

			Celda = Celda & ofv.gencelda("","UT " & wkstylo1,"","","","","",sLink,"si")

			If sActcon <> "" Then
				sImagenP = "&#0232;"
				sLink = OFV.GenLink("", "ft1n", sHref, "", "", sFuncion, sImagenP, "si")
			Else
				sLink="&nbsp;"
			End If
            If session("APLMDCLI") ="BGAL" Then
			Celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sLink,"si")
			end if



			sinput = sinput3

			Celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")

			'sInput =sInput2 'fechaArg(CDate(CLNG(sInput2) + CLng(CDate("01-01-1900"))))' GTFechador(CLNG(sInput2))
			 sInput = GTFechadorSh(CLNG(sInput2) - CLNG(CDATE("01-01-1900")))

			''    response.write " Sinput2b ==> " & Sinput2

			Celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")

			sInput =  "<A HREF='WKFREASIGNADOR.ASP?WKFCOD=" & RS(8) & "&CantRegMoverx=" & CantRegMoverx &"' "

			sInput =  "<A HREF='WKFREASIGNADOR.ASP?ID=" & RS("OBJCOD") & "&Xtip="& Xtip & "&tipobjeto=" & rs("objtipo") & "&WKFCOD=" & RS(8) & "&CantRegMoverx=" & CantRegMoverx &"' "
			sInput =  sInput & "CLASS='NT' style='color:" & Session("CMFPNBOT") & "' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>ASIGNAR</U></A>"
			celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")


			Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")
			posi = posi + 1




			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F9F9F9"
			Else
				xTrColor = "#FFFFFF"
			End If
       end if
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