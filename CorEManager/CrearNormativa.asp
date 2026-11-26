<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->


<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0



	Dim rsprop

	seg=REQUEST.querystring("seg")

	ofv.ObtenerAtributos seg,""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar
	'response.write "STRING: " & request.querystring & "<br>FORM: " & request.form
	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")
	sCodigo = nCodigo
	sacc=request.querystring("pedido")

	idpedido=request.querystring("idpedido")
	SCODFORM=REQUEST.querystring("CODFORM")

	If SCODFORM <> "" Then
		SESSION("SCODFORM")=SCODFORM
	Else
		SCODFORM=SESSION("SCODFORM")
	End If

	sDOCMODAL=request.querystring("docmodal")

	If sDOCMODAL <> "" Then
		session("ModMain")=sDOCMODAL
	Else
		sDOCMODAL=session("ModMain")
	End If
	' RESPONSE.WRITE sDOCMODAL
	CantRegMostrarx = 11
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	Set cn = ofv.conectar(ofv.strconn2)
	Set cn5 = ofv.conectar(ofv.strconn5)

	sHTML = ofv.FormHeader(seg)


	'SHTML =REPLACE(sHTML,"<body","<body beforeunload=""alert('3333');""")
	cerrarventana=""

	DHOY = CLng(Date()) - CLng(CDate("01-01-1900"))

	CLASEFILA="istt"' bgcolor=#f05f23 style='font-weight:bold;'"
	CLASEFILAT="ut style='font-size:12px;'"

	SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='SN'"
	Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
	DESTINO=""
	If Not RS.EOF Then
		DESTINO=Split(RS("TEXTOIMPRESION"),"-")
		TOBJ=RS(0)
		session("tcliente")=TOBJ
	End If
	Call OFV.CERRARCONSULTA(RS)



	Select Case  sACC

		Case "ALTA"
		EXISTE=True
		SQL ="SELECT * FROM PEDIDOS WHERE OBJID=0" & idpedido
		'RESPONSE.WRITE SQL
		Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
		If RS.EOF Then

			SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='SN'"
			Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			RESPONSABLE=""
			If Not RS.EOF Then RESPONSABLE=""'RS("WKFAPRO")
			Call OFV.CERRARCONSULTA(RS)

			sql=" insert into PEDIDOS (SOLICITANTE,ESTADO,FECHAINI,ACCION,TIPOBJETO,responsable,objun) values ('"& SESSION.SESSIONID &"',1," & DHOY & ",'ALTA','" & TOBJ &"','"& RESPONSABLE & "','" & SESSION("CLIEMP") &"')"
			'response.write sql
			Call ofv.crearconsultaEx(sql,cn,1,parametros)
			Strsql = "Select * From PEDIDOS WHERE SOLICITANTE='" & SESSION.SESSIONID & "'"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			IDPEDIDO=RS(0)
			Call OFV.CERRARCONSULTA(RS)

			sql=" UPDATE  PEDIDOS SET SOLICITANTE='" & SESSION("USUARIO") & "' WHERE OBJID="  & IDPEDIDO
			' response.write sql & "<br>"
			Call ofv.crearconsultaEx(sql,cn,1,parametros)

			sql=" insert into PEDIDOSDET (IDPEDIDO,TIPOBJETO) values ('" & IDPEDIDO & "','" & session("tcliente") & "')"
			' response.write sql & "<br>"
			Call ofv.crearconsultaEx(sql,cn,1,parametros)


			If IsArray(DESTINO) Then
				If  UBound(DESTINO)=1 Then
					SQL="INSERT INTO RELAOBJ(CODOBJMASTER,TIPOBJMASTER,CODOJASOC,TIPOBJASOC) VALUES (" & DESTINO(1) & ",'" & DESTINO(0) & "'," & IDPEDIDO & ",'"& TOBJ & "')"
					Call ofv.crearconsultaEx(sql,cn,1,parametros)
				End If
			End If


		End If
		Call OFV.CERRARCONSULTA(RS)

		Case "OK"



		RESPONSE.REDIRECT "SENDWKFSTART.ASP?IDPEDIDO=" & IDPEDIDO & "&TIPOBJETO=" & sTipo


		Case "ELIM"

		SQL="DELETE FROM PEDIDOS WHERE OBJID =0" & IDPEDIDO
		Call ofv.crearconsultaEx(sql,cn,1,parametros)
		SQL="DELETE FROM PEDIDOSDET WHERE idPEDIDO =0" & IDPEDIDO
		Call ofv.crearconsultaEx(sql,cn,1,parametros)
		SQL="DELETE FROM PEDIDOSARCH WHERE idPEDIDO =0" & IDPEDIDO
		Call ofv.crearconsultaEx(sql,cn,1,parametros)
		SQL="DELETE FROM relaobj WHERE codojasoc =0" & IDPEDIDO & " AND tipoBJaSoc='" & TOBJ & "'"
		Call ofv.crearconsultaEx(sql,cn,1,parametros)
		cerrarventana="<SCRIPT>window.top.close();</SCRIPT>"

		'*****************************************
		'********** INCIDENTE 457 ****************
		'*****************************************

		Case "ELIM2"

		'1: MOVER DE CARPETA EL PEDIDO Y CAMBIAR EL ESTADO

		'1.1 Select para conocer la carpeta de destino

		StrCarpDest = "SELECT * FROM tipobjeto "
		StrCarpDest = StrCarpDest & "WHERE tipocpt = 'SN' "
		'RESPONSE.WRITE StrCarpDest & "<BR>"

		Set RsCarpDest = ofv.crearconsultaEx(StrCarpDest,cn,1,parametros)

		'1.1.1	Variable para definir la carpeta de DESTINO

		Dim CarpElim
		CarpElim = ""

		If Not RsCarpDest.EOF Then
			CarpElim = Split(RsCarpDest("Carp_Elim"), "-")
		End If

		'RESPONSE.WRITE CarpElim(0) & "<BR>" & CarpElim(1)



		'1.2: UPDATE de la tabla relaobj para mover el pedido a la carpeta PEDIDOS ELIMINADOS

		SqlUpdDest = "UPDATE relaobj SET codobjmaster = " & CarpElim(1) & " WHERE codojasoc = " & IDPEDIDO & " AND tipoBJaSoc = '" & tobj & "' "
		'RESPONSE.WRITE SqlUpdDest & "<BR>"

		Set RsUpdDest = ofv.crearconsultaEx(SqlUpdDest,cn,1,parametros)


		'1.3 UPDATE del estado del pedido al estado BAJA

		SqlUpdEst = "UPDATE pedidos SET estado = '" & 7 & "' WHERE objid = " & IDPEDIDO & " AND tipobjeto = '" & tobj & "'"
		'RESPONSE.WRITE SqlUpdEst

		Set RsUpdEst = ofv.crearconsultaEx(SqlUpdEst,cn,1,parametros)



		'****************
		'2: INSERTAR Y COMPLETAR LOS PASOS DE WORKFLOW PENDIENTES
		'	INSERTAR EL PASO 999 FINAL
		'***************

		'************
		'2.1 COMPLETAR EL PASO EXISTENTE
		'************

		'2.1.1	Definir una variable para el tipo de WKF
		'		Utilizamos el resultado del recordset RsCarpDest

		Dim TipWKF
		TipWKF = RsCarpDest(19)


		'RESPONSE.WRITE TipWKF & "<BR><BR>"

		'2.1.2	SELECT para Obtener el numero de WKF
		'		Utilizamos la tabla SEGCIRCUITO

		StrWKF = "SELECT segckt FROM SEGCIRCUITO "
		StrWKF = StrWKF & "WHERE objtipo = '" & tobj & "' "
		StrWKF = StrWKF & "AND objcod = '" & Idpedido & "'"
		'response.write StrWKF & "<br><br>"

		Set RsWKF = ofv.crearconsultaEx(StrWKF, cn5, 1, parametros)

		'2.1.3	Definir una VARIABLE para el numero de WKF
		'		Utilizamos el RECORDSET RsWKF

		Dim WkfCkt
		WkfCkt = RsWKF(0)


		'RESPONSE.WRITE "2.1.3 WkfCkt: " & WkfCkt & "<BR>"

		'2.1.4	SELECT para identificar el paso actual y el usuario responsable del paso
		'		Utilizamos la tabla SEGPASO

		StrPasAc = "SELECT * FROM SEGPASO "
		StrPasAc = StrPasAc & "WHERE segckt = '" & WkfCkt & "' "
		StrPasAc = StrPasAc & "AND fin = 'NO' AND okrch = 'NO'"

		'RESPONSE.WRITE StrPasAc & "<br><br>"

		Set RSPasAc = ofv.crearconsultaEx(StrPasAc, cn5, 1, parametros)

		'2.1.5	VARIABLES para identificar el paso actual y el usuario actual
		'		Se va a utilizar para actualizar el paso 999 con este usuario

		Dim PasAc
		Dim UsrAc

		If Not RSPasAc.EOF Then
			'PasAc = RSPasAc(3)
			RespT = RSPasAc(7)
			UsrAc = RSPasAc(8)
			NivAc = RSPasAc(12)
			SubCod = RSPasAc(13)
		End If

		'response.write "2.1.5 PasAc: " & PasAc & "<BR>"
		'response.write PasAc & UsrAc & NivAc & SubCod & "<BR>"

		'2.1.6	Función para armar el paso OKRCH

		If Not RSPasAc.EOF Then
			SEGCKTPSCSIG = GTCKTPSC() & "20"	'Variable que llama a la funcion GTCKTPSC() + campo de orden
		End If

		'RESPONSE.WRITE "2.1.6 SEGCKTPSCSIG: " & SEGCKTPSCSIG & "<br>"

		'2.1.7	VARIABLE Observ1

		Observ1 = "SOLICITUD CANCELADA"

		'2.1.8	Funcion GETDH para completar fecha y hora en los campos correspondientes

		Function GETDH()

			XKD = datepart("d",DATE())
			XKM = datepart("m",DATE())
			XKY = datepart("yyyy",DATE())

			GETDH = XKD & "/" & XKM & "/" & XKY & " " & TIME()

		End Function



		'2.1.9 	Los campos a completar son OKRCH, FIN, OBSERVACIONES

		'		OKRCH: es el campo que identifica al nuevo paso de WKF
		'		FIN: identifica fecha y hora actual de cierre del paso
		'		OBSERVACIONES: se setea una descripcion general en la variable Observ1
		'		Hacemos los 3 UPDATES juntos


		'IF NOT RSPasAc.EOF THEN

		SqlUpdPasAc = "UPDATE segpaso SET okrch = '" & SEGCKTPSCSIG & "', observaciones = '" & Observ1 & "', fin = '" & GETDH & "' "
		SqlUpdPasAc = SqlUpdPasAc & "WHERE segckt = '" & WkfCkt & "' AND fin = 'NO' AND okrch = 'NO' "
		'SqlUpdPasAc = SqlUpdPasAc & "AND model = '" & PasAc & "'"

		Set RSuUpdPasAc = ofv.crearconsultaEx(SqlUpdPasAc, cn5, 1, parametros)
		'response.write SqlUpdPasAc & "<BR>"

		'END IF

		'********************
		'2.2:	INSERTAR EL/LOS PASO/S SIGUIENTE/S
		'********************

		'2.2.1	Select para sacar datos de la tabla wkfcktmaster

		'SQL4 = "SELECT orden, pdescri, wkfpascod  "
		'SQL4 = SQL4 & "FROM wkfcktmaster "
		'SQL4 = SQL4 & "WHERE wkfCOD = '" & TipWKF & "'"
		'SQL4 = SQL4 & "AND WKFPASCOD = '" & PasAc & "'"
		'RESPONSE.WRITE SQL4

		'Set RS41 = ofv.crearconsultaEx(SQL4,cn5,1,parametros)

		'2.2.2	Definicion de variables para hacer el insert de la descripcion del paso en la tabla segpsao
		'		Utilizarmos orden: descriPCION

		'IF NOT RS41.EOF THEN
		'Dim PscDsc
		'PscDsc = RS41(0) & ": " & RS41(1)
		'END IF

		'response.write "2.2.2 PscDsc: " & PscDsc & "<br>"

		'2.2.3	 Función para obtener un nuevo paso siguiente

		'IF NOT rs41.eof THEN
		'SEGCKTPSCSIG2 = GTCKTPSC() & RS41(0) 	'Variable que llama a la funcion GTCKTPSC() + campo de orden
		'END IF

		'response.write "2.2.3 SEGCKTPSCSIG2: " & SEGCKTPSCSIG2 & "<br>"

		'2.2.4	Insertar los nuevos pasos

		'Do WHILE not rs41.eof 'Hacer mientras traiga registros

		'SqlInsPas1 = "INSERT INTO segpaso SELECT segckt, '" & SEGCKTPSCSIG2 & "', segcktpsc, '" & GETDH() & "', '" & GETDH() & "', '" & SEGCKTPSCSIG2 & "', 'USR', codigo, codigo, '" & PscDsc & "', '" & oBSERv1 & "', nivel, subcodigo, null, ID, null from segcktresp where segckt = '" & WkfCkt & "' "
		'SqlInsPas1 = SqlInsPas1 & "AND codigo not in (select responsable from segpaso where SEGCKT = '" & WkfCkt & "') "
		'SqlInsPas1 = SqlInsPas1 & "AND modo = '" & "L'"

		'RESPONSE.WRITE SqlInsPas1

		'SET RsInsPas1 = ofv.crearconsultaEx(SqlInsPas1,cn5,1,parametros)

		'rs41.MOVENEXT

		'LOOP


		'RESPONSE.WRITE SqlIns1

		'**************
		'2.3	INSERT PASO FINAL 999
		'**************

		'2.3.1	Select para sacar el campo PscDsc y el campo model

		StrPasFinal = "SELECT orden, pdescri, wkfpascod "
		StrPasFinal = StrPasFinal & "FROM wkfcktmaster "
		StrPasFinal = StrPasFinal & "WHERE wkfCOD = '" & TipWKF & "'"
		StrPasFinal = StrPasFinal & "AND orden = '999' "

		'RESPONSE.WRITE StrPasFinal & "<BR>"

		Set RS71 = ofv.crearconsultaEx(StrPasFinal,cn5,1,parametros)

		'2.3.2	Definir VARIABLE Para la descripcion del paso y el modelo del paso

		Dim PscDsc2
		PscDsc2 = RS71(0) & ": " & RS71(1)
		Model2 = RS71(2)

		'response.write "2.3.2 PscDsc2: " & PscDsc2 & "<br>"
		'response.write "2.3.2 Model2: " & Model2 & "<br>"

		'2.3.3	Insertar el paso final

		Do While Not RS71.eof 'Hacer mientras traiga registros

			'SqlInsPas2 = "INSERT INTO segpaso SELECT segckt, '" & SEGCKTPSCSIG & "', '" & Model2 & "', '" & GETDH() & "', '" & GETDH() & "', 'NO', 'USR', '" & UsrAc & "', '" & UsrAc & "', '" & PscDsc2 & "', '" & oBSERv1 & "', nivel, subcodigo, null, ID, null from segcktresp where segckt = '" & WkfCkt & "' "
			'SqlInsPas2 = SqlInsPas2 & "and modo = '" & "D'"

			SqlInsFinal = "INSERT INTO SEGPASO (SEGCKT,SEGCKTPSC, MODEL, COMIENZO, FIN, OKRCH, RESPTIPO, RESPONSABLE, RESPASG, PSCDSC, OBSERVACIONES, NIVEL, SUBCODIGO, SEGCKTPSCEXT, SEGIDRESP, SEGSALIDA) VALUES (" & "'" & WkfCkt & "', '" &  SEGCKTPSCSIG & "', '" &  Model2 & "', '" & GETDH() & "', '" & GETDH() & "', 'NO', '" & RespT & "', '" & UsrAc & "', '" & UsrAc & "', '" & PscDsc2 & "', '" & Observ1 & "', " & NivAc & ", " & SubCod & ", " & "NULL, 0, NULL )"

			'RESPONSE.WRITE SqlInsFinal

			Set RsInsPas2 = ofv.crearconsultaEx(SqlInsFinal,cn5,1,parametros)

			RS71.MOVENEXT
		Loop




		cerrarventana="<SCRIPT>window.top.close();</SCRIPT>"

		'******************************
		'****** FIN INCIDENTE 457 *******
		'*******************************



		Case "BAJAF"
		TIPOBJETO=REQUEST.QUERYSTRING("TIPOBJETO")
		CODOBJETO=REQUEST.QUERYSTRING("CODOBJETO")
		' SQL = " UPDATE OBJPRTY SET ESTADO=4 "
		' SQL= SQL & " WHERE CODOBJETO=" & CODOBJETO & " AND TIPOBJETO='" & TIPOBJETO & "'"
		' call ofv.crearconsultaEx(sql,cn,1,parametros)
		' cerrarventana="<SCRIPT>window.top.close();</SCRIPT>"

		response.REDIRECT "asp/EliminarEsquema.asp?IDPEDIDO=" & IDPEDIDO & "&acc=MODIF&Codobjeto=" & CODOBJETO & "&Tipobjeto=" & TIPOBJETO & "&Eliminar=NSNC"

		Case "MODIF"
		TIPOBJETO=REQUEST.QUERYSTRING("TIPOBJETO")
		CODOBJETO=REQUEST.QUERYSTRING("CODOBJETO")

		SQL= "SELECT * FROM tipobjeto WHERE tipobjeto='" & TIPOBJETO & "'"

		'  response.write sql
		Set RSf= ofv.crearconsultaEx(sql,cn,1,parametros)
		sperfil=""
		If Not rsf.eof Then sPerfil=rsf("perfil")
		Call OFV.CERRARCONSULTA(RSf)

		response.REDIRECT "Copiar.asp?IDPEDIDO=" & IDPEDIDO & "&acc=MODIF&Codigo=" & CODOBJETO & "&Tipo=" & TIPOBJETO & "&Perfil=" & sPerfil



		Case "MODI"

		SQL ="SELECT * FROM OBJPROPIEDADES  A "
		SQL= SQL & " INNER JOIN OBJETOS B ON(A.CODOBJETO=B.CODOBJETO )"
		SQL= SQL & " INNER JOIN objhtm B2 ON(A.CODOBJETO=B2.CODOBJETO ) "
		SQL= SQL & " INNER JOIN objprty B1 ON(A.CODOBJETO=B1.CODOBJETO and b1.estado=3) "
		'SQL= SQL & " LEFT JOIN PEDIDOS B3 ON(B3.CODOBJETOL=B.CODOBJETO and b3.estado <> 1 AND B3.ACCION <> 'MODI' ) "
		SQL= SQL & " WHERE  A.CODIGOFORM='" & SCODFORM & "'"
		'response.write sql
		Set RSf= ofv.crearconsultaEx(sql,cn,1,parametros)
		existe=False
		SQL ="SELECT * FROM PEDIDOS WHERE OBJID=0" & idpedido
		'			   RESPONSE.WRITE  "<BR>"  & SQL & "<BR>"
		Set RSDEC = ofv.crearconsultaEx(sql,cn,1,parametros)
		'response.write  RSf.EOF &  " and "  &  RSDEC.eof
		If Not RSf.EOF And RSDEC.eof Then
			existe=True
			SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
			'response.write sql
			Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			RESPONSABLE=""
			If Not RS.EOF Then RESPONSABLE=RS("WKFAPRO")
			Call OFV.CERRARCONSULTA(RS)
			SQL ="SELECT * FROM PEDIDOS WHERE OBJID=0" & idpedido
			' RESPONSE.WRITE  "<BR>"  & SQL & "<BR>"
			Set RSDEC = ofv.crearconsultaEx(sql,cn,1,parametros)


			sql=" insert into PEDIDOS (SOLICITANTE,ESTADO,FECHAINI,CODOBJETOL,TIPOBJETOL,ACCION,RESPONSABLE,tipobjeto,objun) values ('"& SESSION.SESSIONID &"',1," & DHOY & "," & RSF("CODOBJETO") & ",'" & RSF("TIPOBJETO")&"','MODI','" & RESPONSABLE &"','" & TOBJ & "','" & SESSION("CLIEMP") &"')"
			Call ofv.crearconsultaEx(sql,cn,1,parametros)
			Strsql = "Select * From PEDIDOS WHERE SOLICITANTE='" & SESSION.SESSIONID & "'"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			IDPEDIDO=RS(0)
			Call OFV.CERRARCONSULTA(RS)







			strsql="select * from objhtm where codobjeto=" & RSF("CODOBJETO")
			'    RESPONSE.WRITE STRSQL
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			Do While Not rs.eof
				sql="insert into pedidosarch (idpedido,pos,descripcion,archivos) values("& IDPEDIDO & "," & rs("secuen") & ",'" & rs("archivo") & "','')"
				rs.movenext
			Loop



			Call OFV.CERRARCONSULTA(RS)



			sql=" UPDATE  PEDIDOS SET SOLICITANTE='" & SESSION("USUARIO") & "' WHERE OBJID="  & IDPEDIDO
			' response.write sql & "<br>"
			Call ofv.crearconsultaEx(sql,cn,1,parametros)
			COLS="IDPEDIDO"
			DELIM=","
			For I=3 To 14
				If Not IsNull(RSF(I)) Then
					COLS=COLS & DELIM & RSF(I).NAME
				End If
			Next

			COLS=COLS &",NOMBRE,PRODUCTO"
			VALORES= "'" & IDPEDIDO & "'"
			DELIM=","
			For I=3 To 14
				If Not IsNull(RSF(I)) Then
					VALORES=VALORES & DELIM &  "'" & RSF(I) & "'"
				End If
			Next

			VALORES=VALORES & ",'" & RSF(31) & "','" & RSF("TSEGMENTO") & "'"
			sql=" insert into PEDIDOSDET (" & COLS & ") values ("& VALORES & ")"
			' response.write sql & "<br>"
			Call ofv.crearconsultaEx(sql,cn,1,parametros)
			SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
			Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			DESTINO=""
			If Not RS.EOF Then DESTINO=Split(RS("TEXTOIMPRESION"),"-")

			If IsArray(DESTINO) Then
				If  UBound(DESTINO)=1 Then
					SQL="INSERT INTO RELAOBJ(CODOBJMASTER,TIPOBJMASTER,CODOJASOC,TIPOBJASOC) VALUES (" & DESTINO(1) & ",'" &DESTINO(0) & "'," & IDPEDIDO & ",'"& RS(0)& "')"
				End If
			End If

			Call OFV.CERRARCONSULTA(RS)
			Call OFV.CERRARCONSULTA(RSDEC)
			Call ofv.crearconsultaEx(sql,cn,1,parametros)
		End If



		Case "BAJA"

		SQL ="SELECT * FROM OBJPROPIEDADES  A "
		SQL= SQL & " INNER JOIN objprty B1 ON(A.CODOBJETO=B1.CODOBJETO and b1.estado=3) "
		SQL= SQL & " INNER JOIN objhtm B2 ON(A.CODOBJETO=B2.CODOBJETO ) "
		SQL= SQL & " INNER JOIN OBJETOS B ON(A.CODOBJETO=B.CODOBJETO)  WHERE  A.CODIGOFORM='" & SCODFORM & "'"
		'  RESPONSE.WRITE SQL
		Set RSf= ofv.crearconsultaEx(sql,cn,1,parametros)
		existe=False
		If Not RSf.EOF Then
			existe=True
			SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
			Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			RESPONSABLE=""
			DESTINO=""
			If Not RS.EOF Then
				RESPONSABLE=RS("WKFAPRO")
				DESTINO=Split(RS("TEXTOIMPRESION"),"-")
				tobj=rs(0)
				CODOBJETO=RSF("CODOBJETO")
			End If

			Call OFV.CERRARCONSULTA(RS)
			sql=" insert into PEDIDOS (SOLICITANTE,ESTADO,FECHAINI,CODOBJETOL,TIPOBJETOL,ACCION,RESPONSABLE,TIPOBJETO,objun) values ('"& SESSION.SESSIONID &"',1," & DHOY & "," & RSF("CODOBJETO") & ",'" & RSF("TIPOBJETO")&"','ELIM','" & RESPONSABLE &"','"& tobj &"','" & SESSION("CLIEMP") &"')"
			'	 response.write sql & "<br>"

			Call ofv.crearconsultaEx(sql,cn,1,parametros)
			Strsql = "Select * From PEDIDOS WHERE SOLICITANTE='" & SESSION.SESSIONID & "'"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			IDPEDIDO=RS(0)
			Call OFV.CERRARCONSULTA(RS)
			sql=" UPDATE  PEDIDOS SET SOLICITANTE='" & SESSION("USUARIO") & "' WHERE OBJID="  & IDPEDIDO
			' response.write sql & "<br>"
			Call ofv.crearconsultaEx(sql,cn,1,parametros)
			COLS="IDPEDIDO"
			DELIM=","
			For I=3 To 14
				If Not IsNull(RSF(I)) Then
					COLS=COLS & DELIM & RSF(I).NAME
				End If
			Next

			COLS=COLS &",NOMBRE,PRODUCTO"
			VALORES= "'" & IDPEDIDO & "'"
			DELIM=","
			For I=3 To 14
				If Not IsNull(RSF(I)) Then
					VALORES=VALORES & DELIM &  "'" & RSF(I) & "'"
				End If
			Next

			VALORES=VALORES & ",'" & RSF("DESCRI") & "','" & RSF("TSEGMENTO") & "'"
			sql=" insert into PEDIDOSDET (" & COLS & ") values ("& VALORES & ")"
			' response.write sql & "<br>"
			Call OFV.CERRARCONSULTA(RSf)
			Call ofv.crearconsultaEx(sql,cn,1,parametros)


			If IsArray(DESTINO) Then
				If  UBound(DESTINO)=1 Then
					SQL="INSERT INTO RELAOBJ(CODOBJMASTER,TIPOBJMASTER,CODOJASOC,TIPOBJASOC) VALUES (" & DESTINO(1) & ",'" & DESTINO(0) & "'," & IDPEDIDO & ",'"& tobj & "')"
					Call ofv.crearconsultaEx(sql,cn,1,parametros)
				End If
			End If


		End If

		Case "VER"

		XPC=REQUEST.QUERYSTRING("XPC")
		If XPC &"" <> "" Then
			IDPEDIDO=XPC
		Else
			IDPEDIDO= session("ncliente")
		End If

		existe=True
		response.write sdocmodal


		strsql = " select a.estado,A.SOLICITANTE from pedidos a"
		strsql= strsql & "   where a.objid=0" & IDPEDIDO

		'response.write strsql & "<BR>"

		Set rsprop=ofv.crearconsultaEx(StrSql,cn,1,parametros)
		'sdocmodal="WKPED"

		If sdocmodal="WKPED" Then


			tabla1 =  "<div id='DTTEMAINT' name='DTTEMAINT' style='overflow:auto;display:block;"
			tabla1 = TABLA1 & "border-style:solid;border-width:1;border-color:#dddddd;' ></div>" & Chr(13)

			WSPFRMN = "DTTEMAIT"
			'WSPFRMA = "consdocumindex.asp"
			'  WSPFRMA = "CONSTEMAIND.asp?FORMATO=ABR"

			'response.write session("ncliente") & "  " &  session("tcliente")


			WSPDTTSCP = ""
			WSPDTTSCP = WSPDTTSCP & "<script   >"
			' WSPDTTSCP = WSPDTTSCP &  " alert('ENTRO: ');"
			'  WSPDTTSCP = WSPDTTSCP &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
			WSPDTTSCP = WSPDTTSCP & "document." & WSPFRMN & ".submit();"

			If Not rsprop.eof Then

				If rsprop("estado") = 3 Then

					sql = "select objid from objwfldet where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCLiente") & "'"
					Set rdes = ofv.crearconsultaEx(sql,cn,1,parametros)

					If rdes.eof Then

						WSPFRMA="cowkfselect.asp?btnapr=N&acc=ABR&SKIND=1&Tipobjeto=" & session("tcliente") &"&Codobjeto=" & session("ncliente")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
						WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET=WK4 style='display:none;' METHOD='POST' ></FORM>"

					Else

						WSPFRMA="cowkfselect.asp?acc=ABR&SKIND=1&Tipobjeto=" & session("tcliente") &"&Codobjeto=" & session("ncliente")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
						WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET=WK4 style='display:none;' METHOD='POST' ></FORM>"

					End If

				Else

					WSPFRMA="cowkfselect.asp?acc=ABR&SKIND=1&Tipobjeto=" & session("tcliente") &"&Codobjeto=" & session("ncliente")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
					WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET=WK4 style='display:none;' METHOD='POST' ></FORM>"

				End If

			End If
			'''  WSPDTTSCP = WSPDTTSCP &  "document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").style.display="  & chr(34) & "block" & chr(34) & ";"
			WSPDTTSCP = WSPDTTSCP & "</script>"
		End If
	End Select




	strsql = " select b.* ,a.accion,A.CODOBJETOL,A.TIPOBJETOL,a.objid as objidO,a.tipobjeto as tipobjetoO,a.estado as estadoP,A.SOLICITANTE from pedidos a"
	strsql=strsql & " inner join pedidosdet b on (a.objid=b.idpedido)"
	strsql= strsql & "   where b.idPEDIDO=0" & IDPEDIDO

	'response.write strsql & "<BR>"

	Set rsprop=ofv.crearconsultaEx(StrSql,cn,1,parametros)




	sImagenEM  = "<FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>&#0252;</FONT> "
	sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>4</FONT></B> "

	sHTML = sHTML &  "<div id='mt2' >"
	sForm = "edPropiedades"


	If Not rsprop.eof Then
		sEst = rsprop("estadoP")
		existe=True
		codigoform=rsprop("codigoform")
		h1 = ofv.GenCelda("",CLASEFILA,"","10%","","","","Nro Solicitud: ","si")
		h1 = h1 & ofv.GenCelda("","","","45%","","","",rsprop("IDPEDIDO"),"si")


		h1 = h1 & ofv.GenCelda("",CLASEFILA,"","5%","","","","Banca: ","si")

		SDATO="---"

		SDATO=rsprop("BANCA")
		Strsql = "Select * From FORMADMBANCA where estado=3 Order By descriPCION"
		sInput = ofv.GenerarCombo(sForm ,"BANCA",SDATO,Strsql,cn,0,2,"")

		h1 =  h1 &  ofv.GenCelda("","","","45%","","","",sInput,"si")
		FNSHTMLP =   ofv.GenRow("","","","","","","",h1,"si")

		h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","","Fecha Requerida Publicacion: ","si")

		sInputx = ofv.GenerarInput(rsPROP(22).name,rsPROP(22).value,"hidden",sForm,"50","dt","no")

		If isnumeric(rsPROP(22)) Then


			dFecha = cdate(rsPROP(22) + clng(cdate("01-01-1900")))
			xfecact = rsPROP(22)
			'   dFechan = Fechador(xfecact)
			dFechan = GTFechador(rsPROP(22))

		Else
			dFecha = ""
			xfecact = "N"
			dFechan = "-------"

		End If

		sImagenI4 = "<B><FONT FACE=Corflaticon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"
		sinput = "<a class=cc style='text-decoration:none;background-color:none;' xonmouseover='this.style.textDecoration=" & chr(34)
		sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
		sinput = sinput  & "#f7aa22" & chr(34) & ";' xonmouseout='this.style.backgroundColor=" & chr(34)
		sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
		sinput = sinput  & "none" & chr(34) & ";' "
		If session("base")="K"  Then
			sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rsPROP(22).name,GTCSMT,GTCFCH,GTCFCHDAT,"",21,2,0,0,0)
		Else
			sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rsPROP(22).name,GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
		End If
		sinput = sinput  & "  TITLE='Modificar fecha de vigencia.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"


		sInput = dFechan & "&nbsp;" & sinput & sinputx

		h1 =  h1 &  ofv.GenCelda("","","","45%","","","",sInput,"si")

		h1 =  h1 &  ofv.GenCelda("",CLASEFILA,"","5%","","","","Motivo","si")
		sdato="---"
		Strsql = "Select * From motivos order by descripcion"
		sdato = ofv.GenerarCombo(sForm ,rsPROP(20).name,rsPROP(20).value,Strsql,cn,0,1,"")



		h1 =  h1 &  ofv.GenCelda("","","","45%","","","",sdato,"si")


		FNSHTMLP =  FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

		sdato = "---"


		h1 =  ofv.GenCelda("",CLASEFILA,"","5%","","","","Publicacion en WorkPlace:","si")

		sCOMK = "NO" & Chr(9) & "SI"
		sCOMD = "NO" & Chr(9) & "SI"

		sInput = ofv.gencomboFX(sForm,rsPROP(23).name,rsPROP(23).value,scomK,scomD,"si")

		h1 =h1 & ofv.GenCelda("","" ,"","40%","","","",sInput,"si")

		sdato = "---"
		h1 = h1 & ofv.GenCelda("",CLASEFILA,"","5%","","","", "Proceso","si")
		sdato = "---"

		Strsql = "Select * From FORMADMPROCESO where estado=3 Order By descriPCION"
		SDATO = ofv.GenerarCombo(sForm ,rsPROP(12).name,rsPROP(12).value,Strsql,cn,0,2,"")

		h1 = h1 & ofv.GenCelda("","","","45%","","","",sdato,"si")



		FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

		sdato = "---"
		h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Asociado a Proyecto","si")

		SCOMSEP = Chr(9)
		scomK = "SI" & SCOMSEP & "NO"
		scomD = "SI" & SCOMSEP & "NO"

		sInput = ofv.gencomboFX(sForm,"proyecto",rsprop("proyecto"),scomK,scomD,"")

		sInputD=""
		If rsprop("proyecto")="SI" Then
			sInputD = "<BR><textarea class=dt name=DESCRIPROYECTO id=DESCRIPROYECTO rows=4 cols=110 "
			If bModificar  Then
				sInputD = sInputD & " onblur=" & Chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
				sInputD = sInputD & " extenso'); return false;} else {"
				sInputD = sInputD & " document." & sform & ".submit() ;}" & Chr(34)
			Else
				sInputD = sInputD & " readonly "
			End If
			sInputD = sInputD & " >" &  Trim(rsPROP("DESCRIPROYECTO")&"")  &"</textarea></font>"

		End If


		h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput & sInputD,"si")









		'  sql = "select fechabaja from pedidos where objid=" & idpedido
		'  set rg = ofv.crearconsultaex(sql,cn,1,parametros)
		'
		'
		'	sImagenI = "<B><FONT FACE=Corflaticon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"
		'  	sinputE = "<a class=cc style='text-decoration:none;background-color:none;' Xonmouseover='this.style.textDecoration=" & chr(34)
		'  	sinputE = sinputE  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
		'  	sinputE = sinputE  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
		'  	sinputE = sinputE  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
		'  	sinputE = sinputE  & "none" & chr(34) & ";' "
		'  	sinputE = sinputE  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rg(0).name,GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
		'  	sinputE = sinputE  & "  TITLE='Modificar fecha de vencimiento.'  ><b>&nbsp;" & sImagenI & " " & "&nbsp;</b></a>"
		'
		'
		'    h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Fecha Vto:" & sInputE,"si")
		'
		'
		'    sInputD = ofv.GenerarInput(rg(0).name,rg(0).value,"HIDDEN",sForm,"1","dt","")
		'
		'    sInput = "fecha"
		'
		'    h1 = h1 & ofv.GenCelda("","","","45%","","","",sInputD & sInput,"si")
		FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

		h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Descripcion","si")
		sdato = "---"

		sInput = "<textarea class=dt name=DESCRIPCION id=DESCRIPCION rows=4 cols=110 "
		If bModificar  Then
			sInput = sInput & " onblur=" & Chr(34) & " if(this.value.length==0) {this.value=' '}; if (this.value.length > 2300) {alert('El texto es demasiado"
			sInput = sInput & " extenso'); return false;} else {"
			sInput = sInput & " document." & sform & ".submit() ;}" & Chr(34)
		Else
			sInput = sInput & " readonly "
		End If

		xdsc=Trim(rsPROP("DESCRIPCION") & "")

		If Len(xdsc)=0 Or xdsc="" Then xdsc= " "

		sInput = sInput & " >" &  Replace(xdsc & "","'"," ")  &"</textarea></font>"
		h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput,"si")
		FNSHTMLP2 =  ofv.GenRow("","","","","","","",h1,"si")

		'''
		h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Referentes Aprobadores","si")
		sdato = "---"

		sInput = "<textarea class=dt name=" & rsPROP(25).name & " id=" & rsPROP(25).name & " rows=4 cols=110 "
		If bModificar  Then
			sInput = sInput & " onblur=" & Chr(34) & " if(this.value.length==0) {this.value=' '}; if (this.value.length > 2300) {alert('El texto es demasiado"
			sInput = sInput & " extenso'); return false;} else {"
			sInput = sInput & " document." & sform & ".submit() ;}" & Chr(34)
		Else
			sInput = sInput & " readonly "
		End If

		xdsc=Trim(rsPROP(25) & "")

		If Len(xdsc)=0 Or xdsc="" Then xdsc= " "

		sInput = sInput & " >" &  Replace(xdsc & "","'"," ")  &"</textarea></font>"
		h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput,"si")
		FNSHTMLP2 = FNSHTMLP2 & ofv.GenRow("","","","","","","",h1,"si")

		''
		FNSHTMLP2= ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23; ' ","","98%","","0","0","0","0","","",FNSHTMLP2,"si")
		sacc2=sacc
		If sacc="VER" Or sacc="VERW" Then
			xxx=getinfouser(rsPROP("SOLICITANTE"))
			sacc2=rsPROP("ACCION")
		Else
			xxx=getinfouser(session("usuario"))
		End If

		h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>Solicitante:&nbsp;" & session("SGU_nombre") &"</SPAN>","si")
		'sInput = ofv.GenerarInput(rsPROP(15).name,rsPROP(15),"text",SFORM,"100","dt","")
		'h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput,"si")sacc

		FNSHTMLP0 =  ofv.GenRow("","xx bgcolor='"& Session("CMFMENU") &"'","","","45","","",h1,"si")

		Select Case sacc2

			Case "ALTA"
			h1 = ofv.GenCelda("","vt STYLE='text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>SOLICITUD ACTUALIZACION NORMATIVA</SPAN>" ,"si")
			Case "MODI"
			h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>MODIFICACION DE FORMULARIO</SPAN>" ,"si")
			Case "BAJA"
			h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>BAJA DE FORMULARIO</SPAN>" ,"si")
			Case "ELIM"
			h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>BAJA DE FORMULARIO</SPAN>" ,"si")
		End Select


		FNSHTMLP0 =  FNSHTMLP0 & ofv.GenRow("","xxxx bgcolor='"& Session("CMFMENU") & "'","","","45","","",h1,"si")

		h1x = ofv.GenCelda("","","","98% colspan=2","","","", "&nbsp;" ,"si")
		FNSHTMLP0 =  FNSHTMLP0 & ofv.GenRow("","","","","","","",h1x,"si")
		h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "NOMBRE","si")
		sInput = ofv.GenerarInput(rsPROP(15).name,rsPROP(15),"text",SFORM,"100","dt","no")
		h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput,"si")
		FNSHTMLP0 =  FNSHTMLP0 &  ofv.GenRow("","","","","","","",h1,"si")
		FNSHTMLP0= ofv.GenTabla("","aaa  xstyle='border-style:solid;border-width:1;border-color:#f05f23; ' ","","98%","","0","0","0","0","","",FNSHTMLP0,"si")

		FNSHTMLP = FNSHTMLP0 & "<BR>" & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23; ' ","","98%","","0","0","0","0","","",FNSHTMLP,"si")
		FNSHTMLP = FNSHTMLP  & "<BR>" & FNSHTMLP2
		sInput=""


		'  DESCARGAR

		'if sACC ="BAJA"  THEN
		' Strsql = "Select *  From OBJHTM Where CODOBJETO=" & CODOBJETO
		'ELSE
		Strsql = "Select a.*,b.objun  From PEDIDOSARCH a inner join pedidos b on(b.objid=a.idpedido) Where a.idpedido=" & IDPEDIDO & " Order By a.OBJID"

		'END IF
		' RESPONSE.WRITE  Strsql
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		FNSHTMLA =  ofv.GenTabla("","aaa  style='border-style:solid;border-width:1;border-color:#f05f23;' ","","98%","",00,0,0,5,"","","","no")
		Celda = ofv.GenCelda("",clasefila & "  COLSPAN=2 ALIGN=LEFT","","","","","","<font size=2>ARCHIVOS/IMAGENES DE LA NORMATIVA</font>","si")

		sHref="uploadn.asp?acc=" &SACC2&"&id="&IDPEDIDO&"&direct=pedidos&nomdir=pedidos" & "&seg=" & seg & "&vuel=SOL"
		sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & SHREF & "'); return false;" & Chr(34)
		If sACC <>  "VER" And (SACC="ALTA" Or SACC="MODI") Then
			LINKBSQ = ofv.GenLink("", "istt STYLE='CURSOR: hand;border-style:solid;border-SIZE:1;' ", sHref, "", "", sFuncion, "&nbsp;<B>SUBIR</B>&nbsp;", "si")
		Else
			LINKBSQ="&nbsp;"
		End If

		Celda = CELDA & ofv.GenCelda("",clasefila & "  ALIGN=RIGHT ","","","","","",LINKBSQ,"si")

		FNSHTMLA = FNSHTMLA & ofv.GenRow("","","","","","","",Celda,"si")
		POS=1
		Do While Not rs.EOF

			If xTrColor = "#FFFFFF" Then
				xTrColor = "#FFFFFF" '"#F0F0F0"
			Else
				xTrColor = "#FFFFFF"
			End If

			If bSuprimir And (  sACC <>  "VER") Then
				sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
				sHref = sHref & "&Tabla=PEDIDOSARCH&volver=" & vuelta & "_pedido=" &sACC & "_idpedido=" & rsprop(1) & "_seg=" & seg & "&Eliminar=NSNC"
				Celda = ofv.BotonEliminar(rs(1),sHref)
				celda= Replace(celda,"class="," title='Eliminar' class=")
			Else
				Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
			End If


			'  IF SACC ="BAJA"  THEN
			'       stexto1 = "<A class=ttt  target=_new href='" & RS("DIRECTORIO") & "/"  & RS("ARCHIVO")& "'"
			'       stexto1 = stexto1 & "   >" & REPLACE(RS("DESCRI") &"",CHR(13),"<BR>") & "</a>"
			'	ELSE
			If Session("EMPMUN")  = "S" Then
				stexto1 = "<A class=ttt  target=_new href='emp" & rs("objun") & "/" & rs("archivo")&"'"
			Else
				stexto1 = "<A class=ttt  target=_new href='" &rs("archivo")&"'"
			End If

			'stexto1 = "<A class=ttt  target=_new href='" &rs("archivo")&"'"
			stexto1 = stexto1 & "   >" & Replace(RS("DESCRIPCION") &"",Chr(13),"<BR>") & "</a>"

			'END IF


			Celda = CELDA & ofv.GenCelda("",CLASEFILAT,"","90%","","","", stexto1,"si")




			FNSHTMLA = FNSHTMLA & ofv.GenRow("","ut  bgcolor='" & xTrColor & "'  ","","","","","",Celda,"si")

			RS.MOVENEXT
			POS=POS +1
		Loop

		FNSHTMLA =FNSHTMLA & "</table>"
		' RESPONSE.WRITE sACC
		If bModificar  And SACC <> "BAJA" Then
			sAccion = "asp/ActualizarEN.asp?ID=" & rsprop(0) & "&Tabla=pedidosdet&volver=" & vuelta  & "_idpedido=" & rsprop(1) & "_seg=" & seg   & "_PEDIDO=" & SACC &"_docmodal=" & sdocmodal
			sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
			Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",FNSHTMLP,"si")
		Else
			Formulario =FNSHTMLP
		End If

		sHTML = sHTML & Formulario

		sHTML = sHTML  & FNSHTMLA
		'response.write sACC
		If sEst = 3 And sdocmodal = "WKPED" Then

			sHTML = sHTML & "<br><br>"
			sImage = "<font face='corflaticon' size=3 color='#FE9A2E;'>a</font>"
			sLink = "<a class='cc' onclick=""window.onbeforeunload = null;"" href='edobjwfldet.asp?codobjeto=" & session("NCliente") & "&tipobjeto=" & session("TCliente") & "' >SELECCIONAR DESARROLLADORES</a>"
			sCelda = ofv.GenCelda("","a","center","","","","",sImage & sLink,"si")
			sFila = ofv.GenRow("","","","","","","",sCelda,"si")
			sTabla = ofv.GenTabla("","aaa ","","100% ","","0","","0","0","","",sFila,"si")
			sHTML = sHTML  & sTabla

		End If

		If sACC <>  "VER" And sACC <>  "VERW"   Then

			sForm="ALTA"

			If validardatos(rsprop,SACC) Then
				BTaccadd = " onClick=" & Chr(34) &  "this.style.display='none';document.getElementById('cn').style.display='none'; document.getElementById('info').style.display='';document." & sForm & ".submit();" & Chr(34) &  " "
			Else
				If SACC="ALTA" Then
					BTaccadd = " onClick=" & Chr(34) &  "alert('DEBE COMPLETAR LOS CAMPOS DE NOMBRE, FECHA Y DESCRIPCION PARA CURSAR LA SOLICITUD.');" & Chr(34) &  " "
				ElseIf SACC="MODI" Then
					BTaccadd = " onClick=" & Chr(34) &  "alert('Debe agregar un ATTACH para cursar el pedido.');" & Chr(34) &  " "
				ElseIf SACC="BAJA" Then
					BTaccadd = " onClick=" & Chr(34) &  "alert('Debe agregar un ATTACH para cursar el pedido.');" & Chr(34) &  " "
				Else


					BTaccadd = " onClick=" & Chr(34) &  "alert('Debe completar toda la informacion para cursar el pedido.');" & Chr(34) &  " "
				End If

			End If




			SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CONFIRMAR  ' CLASS='BT' " & BTaccadd & " >"


			SACCION="CrearNormativa.asp?PEDIDO=OK&idpedido=" & rsprop(1) & "&TIPOBJETO=" & tobj
			Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")
			'CONFIRMAR=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",Formulario,"si")
			Celda = ofv.gencelda("","ut","RIGHT","50%","","","",Formulario,"si")

			sForm="BAJA"
			SINPUT = "<INPUT id=cn  name=cn TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"



			SACCION="CrearNormativa.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)
			Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")
			Celda = CELDA & ofv.gencelda("","ut","LEFT","50%","","","",Formulario,"si")
			Filax  =  ofv.genrow("","AAA  " ,"","","","","",Celda,"si")
			sHTML = sHTML & "<br>" & ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")
			If rsprop("estadop") ="4" Or  sDOCMODAL="WKPED" Then TABLA=""

		Else
			sForm="FORMACC"

			' se fija accion del pedido alta,modi elim

			Select Case   rsprop("accion") & ""


				Case "ALTA"

				SACCION="asp/AgregarEsquema.asp?acc=ALTAPED&idpedido=" & idpedido
				SACCION="Copiar2.asp?BASEF=PEDIDO&acc=ALTA&idpedido=" & idpedido

				Select Case RSPROP("ESTADOp")
					Case 3

					'SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CREAR FORMULARIO  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
					Case 1
					'SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
					SACCION="CrearNormativa.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)

				End Select

				If  sACC <>  "VERW" Then

					Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")
				Else
					Formulario=""
				End If

				Celda = ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")



				SQL= "select * from CKTVIEW "
				SQL= SQL & " where  "
				SQL=SQL  & "OBJTIPO='"& rsprop("TIPOBJETOo") &"' and objcod='" &rsprop("objido") & "' and PSCFIN='NO' "
				' response.write sql & sACC
				Set rs5 = ofv.crearconsultaEx(SQL,cn5,1,parametros)

				If Not RS5.EOF And rsprop("estadop") <> 3 Then
					session("SEGCKTPSC")=RS5("SEGCKTPSC")

					saccion="cowkfselect.asp?OPCION=RCH&ACC=CANCELA&SKIND=1&Tipobjeto=" &rsprop("TIPOBJETOo") &"&Codobjeto=" & rsprop("objido")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"

					'SINPUT = "<INPUT TYPE='BUTTON' VALUE=' CANCELA SOLICITUD  ' CLASS='BT' onClick='document." & sForm & "WKF" & ".submit();'>"
					If  sACC <>  "VERW" Then

						Formulario = ofv.GenForm(sForm & "WKF","",sAccion,"","POST","",SINPUT,"si")
					Else
						Formulario =""
					End If

					Celda = CELDA & ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")
				End If
				FILAX = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
				TABLA=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")
				If rsprop("estadop") ="4" Or  sDOCMODAL="WKPED" Then TABLA=""

				Case "MODI"


				'SACCION="Copiar2.asp?BASEF=PEDIDO&acc=ALTA&idpedido=" & idpedido
				SACCION="CrearNormativa.asp?PEDIDO=MODIF&idpedido=" & idpedido & "&CODOBJETO=" & rsprop("CODOBJETOL") & "&TIPOBJETO=" & rsprop("TIPOBJETOL")

				If rsprop("estadop") = 3 Then
					SINPUT = "<INPUT TYPE='BUTTON' VALUE='  MODIFICAR FORMULARIO  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
				ElseIf rsprop("estadop") = 1 Then
					'SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
					SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
					SACCION="CrearNormativa.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)
				Else
					SINPUT=""
				End If



				Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")

				Celda = ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")

				SQL= "select * from CKTVIEW "
				SQL= SQL & " where  "
				SQL=SQL  & "OBJTIPO='"& rsprop("TIPOBJETOo") &"' and objcod='" &rsprop("objido") & "' and PSCFIN='NO' "
				'  response.write sql
				Set rs5 = ofv.crearconsultaEx(SQL,cn5,1,parametros)

				If Not RS5.EOF And rsprop("estadop") <> 3 Then
					session("SEGCKTPSC")=RS5("SEGCKTPSC")

					saccion="cowkfselect.asp?OPCION=RCH&ACC=CANCELA&SKIND=1&Tipobjeto=" &rsprop("TIPOBJETOo") &"&Codobjeto=" & rsprop("objido")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"

					SINPUT = "<INPUT TYPE='BUTTON' VALUE=' CANCELA SOLICITUD  ' CLASS='BT' onClick='document." & sForm & "WKF" & ".submit();'>"

					Formulario = ofv.GenForm(sForm & "WKF","",sAccion,"","POST","",SINPUT,"si")

					Celda = CELDA & ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")
				End If


				FILAX = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
				TABLA=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")

				If rsprop("estadop") ="4" Or  sDOCMODAL="WKPED" Then TABLA=""

				Case "ELIM"


				'SACCION="Copiar2.asp?BASEF=PEDIDO&acc=ALTA&idpedido=" & idpedido
				SACCION="CrearNormativa.asp?PEDIDO=BAJAF&idpedido=" & idpedido & "&CODOBJETO=" & rsprop("CODOBJETOL") & "&TIPOBJETO=" & rsprop("TIPOBJETOL")

				If rsprop("estadop") = 3 Then
					SINPUT = "<INPUT TYPE='BUTTON' VALUE='  BAJA DE  FORMULARIO  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"

				ElseIf rsprop("estadop") = 1 Then
					'SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
					SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
					SACCION="CrearNormativa.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)
				Else
					SINPUT=""

				End If

				Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")

				Celda = ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")

				SQL= "select * from CKTVIEW "
				SQL= SQL & " where  "
				SQL=SQL  & "OBJTIPO='"& rsprop("TIPOBJETOo") &"' and objcod='" &rsprop("objido") & "' and PSCFIN='NO' "
				'  response.write sql
				Set rs5 = ofv.crearconsultaEx(SQL,cn5,1,parametros)

				If Not RS5.EOF And rsprop("estadop") <> 3 Then
					session("SEGCKTPSC")=RS5("SEGCKTPSC")

					saccion="cowkfselect.asp?OPCION=RCH&ACC=CANCELA&SKIND=1&Tipobjeto=" &rsprop("TIPOBJETOo") &"&Codobjeto=" & rsprop("objido")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"

					SINPUT = "<INPUT TYPE='BUTTON' VALUE=' CANCELA SOLICITUD  ' CLASS='BT' onClick='document." & sForm & "WKF" & ".submit();'>"

					Formulario = ofv.GenForm(sForm & "WKF","",sAccion,"","POST","",SINPUT,"si")

					Celda = CELDA & ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")
				End If
				FILAX = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
				TABLA=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")
				If rsprop("estadop") ="4" Or  sDOCMODAL="WKPED" Then TABLA=""



			End Select

		End If

	End If
	sHTML = sHTML & "<br><br> " & TABLA
	sHTML = sHTML & "</div>"

	sInput="<B>GENERANDO  WORKFLOW DE SOLICITUD , AGUARDE POR FAVOR</B>"
	celda =  ofv.GenCelda("","tt style='text-align: center !important;border-style:solid;border-width:1;border-color:#dddddd;'","","","50%","","",sInput,"si")
	fila=ofv.GenRow("","","","","","","",celda,"si")
	TAB =  ofv.GenTabla("","","","98% id=info name=info  style='display:none;'","","0","","0","0","","",fila,"si")






	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","") & TAB



	SHTML=SHTML & cerrarventana  & tabla1 & WSPFORME & WSPDTTSCP & r '& "222"

	If Not existe Then
		shtml = shtml & "<TABLE><TR><TD CLASS=UT><FONT color=#993333 size=3 face=WEBDINGS>i</FONT> NO EXISTE FORMULARIO CON EL CODIGO " & SCODFORM & "</TD></TR>"
		SHREF="PedidoForm.ASP?Ing=P&BASE=K&seg=CLIMN52412655104952828166469"
		sform="<form name=xsf id=xsf method=post action='" & SHREF & "'></form>"
		sFuncion=" onclick='document.xsf.submit();' "
		LINKBSQ2="<tr><td><BR> <input type=button  class=bietx	 xSTYLE='width:90%;height:50px;color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' " & sFuncion & " value='volver' ></td></tr></table>"
		shtml=shtml & LINKBSQ2 & sform
	End If


	If sdocmodal="ED" Then
		xsFormulario = "<form name=nuloz id=nuloz target=fesquemas action=" & Session("ASP") & "  method=post>"
		shtml=SHTML & xsFormulario & "<SCRIPT> "
		shtml=SHTML & " window.top.CorWin.addEventListener('beforeunload', function (e) { "
		shtml=SHTML & "          document.getElementById('nuloz').submit();"
		shtml=SHTML & "          });"
		shtml=SHTML & "</script>"

	End If

	If sdocmodal="WKPED" Then
		xsFormulario = "<form name=nuloz id=nuloz action='novedmenu.asp?tab=view4' target=info method=post>"
		shtml=SHTML & xsFormulario & "<SCRIPT> "
		shtml=SHTML & " window.top.CorWin.addEventListener('beforeunload', function (e) { "
		shtml=SHTML & "          document.getElementById('nuloz').submit();"
		shtml=SHTML & "          });"
		shtml=SHTML & "</script>"




	End If


	If sdocmodal="CL" Then
		'  xsFormulario = "<form name=nuloz id=nuloz action='novedmenu.asp' target=info method=post>"
		shtml=SHTML &  "<SCRIPT> function handleWindowClose(e) { "
		shtml=SHTML & "   e = window.event || e; "
		'  shtml=SHTML & "       if ((e.clientX < 0) || (e.clientY < 0))"
		'   shtml=SHTML & "       {"

		shtml=SHTML & "         e.preventDefault();e.stopImmediatePropagation();e.stopPropagation();"
		'   shtml=SHTML & "        }"
		shtml=SHTML & "}"
		'   shtml=SHTML & " window.onbeforeunload = handleWindowClose;</script>"
	End If

	sHTML = sHTML & "<script> var element=document.getElementById('MTH'); if (typeof(element) != 'undefined' && element != null) {element.style.display='none';}</script>" & GTCCALENDAR("D") & "</body></html>"
	Call ofv.cerrarconn(cn)
	Call ofv.cerrarconn(cn5)


End If


response.write sHTML

Function validardatos(rspropx,ACC)




	'IF ACC="BAJA" THEN
	' validardatos=TRUE
	' EXIT FUNCTION
	'END IF

	validardatos=False


	Select Case  ACC

		Case "BAJA"

		'     SQL="SELECT OBJID FROM PEDIDOSARCH WHERE IDPEDIDO=" & rspropx(1)

		'	 	set rsATT = ofv.crearconsultaEx(sql,cn,1,parametros)

		' 	IF NOT rsATT.EOF THEN
		validardatos=True
		'	END IF
		'	CALL OFV.CERRARCONSULTA(rsATT)


		Exit Function


		Case "ALTA"




		'SQL="SELECT OBJID FROM PEDIDOSARCH WHERE IDPEDIDO=" & rspropx(1)
		' RESPONSE.WRITE SQL
		'set rsATT = ofv.crearconsultaEx(sql,cn,1,parametros)

		If  Trim(rspropx("nombre")& "") = "" Or Trim(rspropx("DESCRIPCION")& "") = "" Or Trim(rspropx("p11")& "") = "" Then
			Exit Function
		End If

		'if  trim(rspropx("DESCRIPCION")& "") = "" then
		'	  exit function
		'end if

		'IF NOT rsATT.EOF THEN
		' validardatos=TRUE
		'END IF
		'CALL OFV.CERRARCONSULTA(rsATT)
		'EXIT FUNCTION
		validardatos=True

		Case "MODI"
		validardatos=True

		SQL="SELECT OBJID FROM PEDIDOSARCH WHERE IDPEDIDO=" & rspropx(1)

		Set rsATT = ofv.crearconsultaEx(sql,cn,1,parametros)
		If   rsATT.EOF Then
			validardatos=False
		End If
		Call OFV.CERRARCONSULTA(rsATT)

		'SQL="SELECT A.* , OBJETOS.TSEGMENTO FROM  objpropiedades  A"
		'SQL=SQL & " INNER JOIN OBJETOS ON (OBJETOS.CODOBJETO=A.CODOBJETO) "
		'SQL=SQL & " WHERE A.codobjeto=" & rspropx("codobjetol")
		'RESPONSE.WRITE SQL
		'set rsATT = ofv.crearconsultaEx(sql,cn,1,parametros)
		'IF NOT rsATT.EOF THEN

		' if  rspropx("banca")& "" <> rsATT("banca") & "" then

		'    EXIT FUNCTION
		' end if



		'if  rspropx("tiempoguarda")& "" <> rsATT("tiempoguarda")& "" then

		'   EXIT FUNCTION
		'end if


		'if  rspropx("catform")& "" <>  rsATT("catform")& "" then

		'  exit function
		'end if


		'if  rspropx("firmasform")& "" <> rsATT("firmasform") & "" then
		'  exit function
		'end if

		'if  rspropx("modoimpresion")& "" <> rsATT("modoimpresion")& ""  then
		'  exit function
		'end if



		'if  rspropx("proceso")& "" <>  rsATT("proceso")& "" then
		'  exit function
		'end if

		'if  rspropx("producto")& "" <> rsATT("TSEGMENTO")& "" then
		'  exit function
		'end if

		''if  rspropx("nombre")& ""  <> rsATT("nombre")& "" then
		' ' exit function
		''end if



		'if  rspropx("motivo")& "" <>  rsATT("motivo")& "" then
		'  exit function
		'end if


		'validardatos=FALSE

		'ELSE


		'validardatos=FALSE
		'END IF



		'END IF

		Case Else


		If  rspropx("banca")& "" ="" Then
			Exit Function
		End If



		If  rspropx("tiempoguarda")& "" ="" Then
			Exit Function
		End If


		If  rspropx("catform")& "" ="" Then
			Exit Function
		End If


		If  rspropx("firmasform")& "" ="" Then
			Exit Function
		End If

		If  rspropx("modoimpresion")& "" ="" Then
			Exit Function
		End If



		If  rspropx("proceso")& "" ="" Then
			Exit Function
		End If

		If  rspropx("producto")& "" ="" Then
			Exit Function
		End If

		If  rspropx("nombre")& "" ="" Then
			Exit Function
		End If

		If  rspropx("publiconectado")& "" ="" Then
			Exit Function
		End If

		If  rspropx("motivo")& "" ="" Then
			Exit Function
		End If

		If  rspropx("tipoformu")& "" ="" Then
			Exit Function
		End If
		validardatos=True
	End Select

	validardatos=True

End Function


%>

