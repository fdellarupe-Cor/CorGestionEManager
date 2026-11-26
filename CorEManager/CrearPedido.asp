<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->


<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
	uv = ofv.uv
	if uv = "NO" then ok = ofv.ValidarUsuario
	sHTML = ok
else
	response.expires=0



	dim rsprop
   ' xsFormulario=""
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

	IF SCODFORM <> "" THEN
	SESSION("SCODFORM")=SCODFORM
	ELSE
	SCODFORM=SESSION("SCODFORM")
	END IF



   sDOCMODAL=request.querystring("docmodal")

   if sDOCMODAL <> "" then
	  session("ModMain")=sDOCMODAL
	else
	  sDOCMODAL=session("ModMain")
	end if
   'RESPONSE.WRITE sDOCMODAL
	CantRegMostrarx = 11
	CantRegMoverx = cdbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	set cn = ofv.conectar(ofv.strconn2)
    set cn5 = ofv.conectar(ofv.strconn5)

	sHTML = ofv.FormHeader(seg)

 'RESPONSE.WRITE REQUEST.querystring
	'SHTML =REPLACE(sHTML,"<body","<body beforeunload=""alert('3333');""")
	cerrarventana=""

	DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))

	CLASEFILA="istt"' bgcolor=#f05f23 style='font-weight:bold;'"
	CLASEFILAT="ut style='font-size:12px;'"

	SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
			set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			DESTINO=""
	IF NOT RS.EOF THEN
		DESTINO=SPLIT(RS("TEXTOIMPRESION"),"-")
		TOBJ=RS(0)
       session("tcliente")=TOBJ
	END IF
	CALL OFV.CERRARCONSULTA(RS)



	SELECT CASE  sACC

	CASE "ALTA"
        EXISTE=TRUE
	   SQL ="SELECT * FROM PEDIDOS WHERE OBJID=0" & idpedido
	   'RESPONSE.WRITE SQL
	   set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
       IF RS.EOF THEN

        	SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
			set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			RESPONSABLE=""
			IF NOT RS.EOF THEN RESPONSABLE=""'RS("WKFAPRO")
			CALL OFV.CERRARCONSULTA(RS)

		sql=" insert into PEDIDOS (SOLICITANTE,ESTADO,FECHAINI,ACCION,TIPOBJETO,responsable,objun) values ('"& SESSION.SESSIONID &"',1," & DHOY & ",'ALTA','" & TOBJ &"','"& RESPONSABLE &"','" & SESSION("CLIEMP") &"')"
        'response.write sql
		call ofv.crearconsultaEx(sql,cn,1,parametros)
		Strsql = "Select * From PEDIDOS WHERE SOLICITANTE='" & SESSION.SESSIONID & "'"
		set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		IDPEDIDO=RS(0)
		CALL OFV.CERRARCONSULTA(RS)

		sql=" UPDATE  PEDIDOS SET SOLICITANTE='" & SESSION("USUARIO") & "' WHERE OBJID="  & IDPEDIDO
' response.write sql & "<br>"
		call ofv.crearconsultaEx(sql,cn,1,parametros)

		sql=" insert into PEDIDOSDET (IDPEDIDO,tipobjeto) values ('" & IDPEDIDO & "','" & session("tcliente") & "')"
' response.write sql & "<br>"
		call ofv.crearconsultaEx(sql,cn,1,parametros)


		IF ISARRAY(DESTINO) THEN
			IF  ubound(DESTINO)=1 THEN
				SQL="INSERT INTO RELAOBJ(CODOBJMASTER,TIPOBJMASTER,CODOJASOC,TIPOBJASOC) VALUES (" & DESTINO(1) & ",'" &DESTINO(0) & "'," & IDPEDIDO & ",'"& TOBJ & "')"

				call ofv.crearconsultaEx(sql,cn,1,parametros)

			END IF
		END IF


   END IF
       CALL OFV.CERRARCONSULTA(RS)

	CASE "OK"



		RESPONSE.REDIRECT "SENDWKFSTART.ASP?IDPEDIDO=" & IDPEDIDO & "&TIPOBJETO=" & sTipo


	CASE "ELIM"

		SQL="DELETE FROM PEDIDOS WHERE OBJID =0" & IDPEDIDO
		call ofv.crearconsultaEx(sql,cn,1,parametros)
		SQL="DELETE FROM PEDIDOSDET WHERE idPEDIDO =0" & IDPEDIDO
		call ofv.crearconsultaEx(sql,cn,1,parametros)
		SQL="DELETE FROM PEDIDOSARCH WHERE idPEDIDO =0" & IDPEDIDO
		call ofv.crearconsultaEx(sql,cn,1,parametros)
		SQL="DELETE FROM relaobj WHERE codojasoc =0" & IDPEDIDO & " AND tipoBJaSoc='" & TOBJ & "'"
		call ofv.crearconsultaEx(sql,cn,1,parametros)
		cerrarventana="<SCRIPT>window.top.close();</SCRIPT>"

	'*****************************************
	'********** INCIDENTE 457 ****************
	'*****************************************

	CASE "ELIM2"

		'1: MOVER DE CARPETA EL PEDIDO Y CAMBIAR EL ESTADO

		'1.1 Select para conocer la carpeta de destino

			StrCarpDest = "SELECT * FROM tipobjeto "
			StrCarpDest = StrCarpDest & "WHERE tipocpt = 'PD' "
			'RESPONSE.WRITE StrCarpDest & "<BR>"

			SET RsCarpDest = ofv.crearconsultaEx(StrCarpDest,cn,1,parametros)

			'1.1.1	Variable para definir la carpeta de DESTINO

			Dim CarpElim
			CarpElim = ""

				IF NOT RsCarpDest.EOF THEN
					CarpElim = Split(RsCarpDest("Carp_Elim"), "-")
				END IF

			'RESPONSE.WRITE CarpElim(0) & "<BR>" & CarpElim(1)



		'1.2: UPDATE de la tabla relaobj para mover el pedido a la carpeta PEDIDOS ELIMINADOS

			SqlUpdDest = "UPDATE relaobj SET codobjmaster = " & CarpElim(1) & " WHERE codojasoc = " & IDPEDIDO & " AND tipoBJaSoc = '" & tobj & "' "
			'RESPONSE.WRITE SqlUpdDest & "<BR>"

			SET RsUpdDest = ofv.crearconsultaEx(SqlUpdDest,cn,1,parametros)


		'1.3 UPDATE del estado del pedido al estado BAJA

			SqlUpdEst = "UPDATE pedidos SET estado = '" & 6 & "' WHERE objid = " & IDPEDIDO & " AND tipobjeto = '" & tobj & "'"
			'RESPONSE.WRITE SqlUpdEst

			SET RsUpdEst = ofv.crearconsultaEx(SqlUpdEst,cn,1,parametros)



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

						SET RsWKF = ofv.crearconsultaEx(StrWKF, cn5, 1, parametros)

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

						SET RSPasAc = ofv.crearconsultaEx(StrPasAc, cn5, 1, parametros)

					'2.1.5	VARIABLES para identificar el paso actual y el usuario actual
					'		Se va a utilizar para actualizar el paso 999 con este usuario

						Dim PasAc
						Dim UsrAc

						IF NOT RSPasAc.EOF THEN
						'PasAc = RSPasAc(3)
						RespT = RSPasAc(7)
						UsrAc = RSPasAc(8)
						NivAc = RSPasAc(12)
						SubCod = RSPasAc(13)
						END IF

						'response.write "2.1.5 PasAc: " & PasAc & "<BR>"
						'response.write PasAc & UsrAc & NivAc & SubCod & "<BR>"

					'2.1.6	Función para armar el paso OKRCH

						IF NOT RSPasAc.EOF THEN
						SEGCKTPSCSIG = GTCKTPSC() & "20"	'Variable que llama a la funcion GTCKTPSC() + campo de orden
						END IF

						'RESPONSE.WRITE "2.1.6 SEGCKTPSCSIG: " & SEGCKTPSCSIG & "<br>"

					'2.1.7	VARIABLE Observ1

						Observ1 = "PEDIDO CANCELADO"

					'2.1.8	Funcion GETDH para completar fecha y hora en los campos correspondientes

						FUNCTION GETDH()

						XKD = datepart("d",DATE())
						XKM = datepart("m",DATE())
						XKY = datepart("yyyy",DATE())

						GETDH = XKD & "/" & XKM & "/" & XKY & " " & TIME()

						END FUNCTION



					'2.1.9 	Los campos a completar son OKRCH, FIN, OBSERVACIONES

					'		OKRCH: es el campo que identifica al nuevo paso de WKF
					'		FIN: identifica fecha y hora actual de cierre del paso
					'		OBSERVACIONES: se setea una descripcion general en la variable Observ1
					'		Hacemos los 3 UPDATES juntos


						'IF NOT RSPasAc.EOF THEN

						SqlUpdPasAc = "UPDATE segpaso SET okrch = '" & SEGCKTPSCSIG & "', observaciones = '" & Observ1 & "', fin = '" & GETDH & "' "
						SqlUpdPasAc = SqlUpdPasAc & "WHERE segckt = '" & WkfCkt & "' AND fin = 'NO' AND okrch = 'NO' "
						'SqlUpdPasAc = SqlUpdPasAc & "AND model = '" & PasAc & "'"

						SET RSuUpdPasAc = ofv.crearconsultaEx(SqlUpdPasAc, cn5, 1, parametros)
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

						SET RS71 = ofv.crearconsultaEx(StrPasFinal,cn5,1,parametros)

					'2.3.2	Definir VARIABLE Para la descripcion del paso y el modelo del paso

						Dim PscDsc2
						PscDsc2 = RS71(0) & ": " & RS71(1)
						Model2 = RS71(2)

						'response.write "2.3.2 PscDsc2: " & PscDsc2 & "<br>"
						'response.write "2.3.2 Model2: " & Model2 & "<br>"

					'2.3.3	Insertar el paso final

						Do WHILE not RS71.eof 'Hacer mientras traiga registros

						'SqlInsPas2 = "INSERT INTO segpaso SELECT segckt, '" & SEGCKTPSCSIG & "', '" & Model2 & "', '" & WkfCkt & "', '" & GETDH() & "', 'NO', 'USR', '" & UsrAc & "', '" & UsrAc & "', '" & PscDsc2 & "', '" & oBSERv1 & "', nivel, subcodigo, null, ID, null from segcktresp where segckt = '" & WkfCkt & "' "
						'SqlInsPas2 = SqlInsPas2 & "and modo = '" & "D'"

						SqlInsFinal = "INSERT INTO SEGPASO (SEGCKT,SEGCKTPSC, MODEL, COMIENZO, FIN, OKRCH, RESPTIPO, RESPONSABLE, RESPASG, PSCDSC, OBSERVACIONES, NIVEL, SUBCODIGO, SEGCKTPSCEXT, SEGIDRESP, SEGSALIDA) VALUES (" & "'" & WkfCkt & "', '" &  SEGCKTPSCSIG & "', '" &  Model2 & "', '" & GETDH() & "', '" & GETDH() & "', 'NO', '" & RespT & "', '" & UsrAc & "', '" & UsrAc & "', '" & PscDsc2 & "', '" & Observ1 & "', " & NivAc & ", " & SubCod & ", " & "NULL, 0, NULL )"

						'RESPONSE.WRITE SqlInsFinal

						SET RsInsPas2 = ofv.crearconsultaEx(SqlInsFinal,cn5,1,parametros)

						RS71.MOVENEXT
						LOOP




		                                                                                                                                                                    cerrarventana="<SCRIPT>window.top.close();</SCRIPT>"

		'******************************
		'****** FIN INCIDENTE 457 *******
		'*******************************



    CASE "BAJAF"
          TIPOBJETO=REQUEST.QUERYSTRING("TIPOBJETO")
          CODOBJETO=REQUEST.QUERYSTRING("CODOBJETO")
         ' SQL = " UPDATE OBJPRTY SET ESTADO=4 "
         ' SQL= SQL & " WHERE CODOBJETO=" & CODOBJETO & " AND TIPOBJETO='" & TIPOBJETO & "'"
	     ' call ofv.crearconsultaEx(sql,cn,1,parametros)
		 ' cerrarventana="<SCRIPT>window.top.close();</SCRIPT>"

		 response.REDIRECT "asp/EliminarEsquema.asp?IDPEDIDO=" & IDPEDIDO & "&acc=MODIF&Codobjeto=" & CODOBJETO & "&Tipobjeto=" & TIPOBJETO & "&Eliminar=NSNC"

    CASE "MODIF"
          TIPOBJETO=REQUEST.QUERYSTRING("TIPOBJETO")
          CODOBJETO=REQUEST.QUERYSTRING("CODOBJETO")

          SQL= "SELECT * FROM tipobjeto WHERE tipobjeto='" & TIPOBJETO & "'"

        '  response.write sql
         SET RSf= ofv.crearconsultaEx(sql,cn,1,parametros)
         sperfil=""
         if not rsf.eof then sPerfil=rsf("perfil")
         CALL OFV.CERRARCONSULTA(RSf)

          response.REDIRECT "Copiar.asp?IDPEDIDO=" & IDPEDIDO & "&acc=MODIF&Codigo=" & CODOBJETO & "&Tipo=" & TIPOBJETO & "&Perfil=" & sPerfil



	CASE "MODI"

		SQL ="SELECT * FROM OBJPROPIEDADES  A "
		SQL= SQL & " INNER JOIN OBJETOS B ON(A.CODOBJETO=B.CODOBJETO )"
		SQL= SQL & " INNER JOIN objhtm B2 ON(A.CODOBJETO=B2.CODOBJETO ) "
		SQL= SQL & " INNER JOIN objprty B1 ON(A.CODOBJETO=B1.CODOBJETO and b1.estado=3) "
		'SQL= SQL & " LEFT JOIN PEDIDOS B3 ON(B3.CODOBJETOL=B.CODOBJETO and b3.estado <> 1 AND B3.ACCION <> 'MODI' ) "
		SQL= SQL & " WHERE  A.CODIGOFORM='" & SCODFORM & "'"
		'response.write sql
		SET RSf= ofv.crearconsultaEx(sql,cn,1,parametros)
        existe=false
        SQL ="SELECT * FROM PEDIDOS WHERE OBJID=0" & idpedido
			'			   RESPONSE.WRITE  "<BR>"  & SQL & "<BR>"
	        set RSDEC = ofv.crearconsultaEx(sql,cn,1,parametros)
			'response.write  RSf.EOF &  " and "  &  RSDEC.eof
			IF NOT RSf.EOF and RSDEC.eof THEN
		    existe=true
			SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
			'response.write sql
			set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			RESPONSABLE=""
			IF NOT RS.EOF THEN RESPONSABLE=RS("WKFAPRO")
			CALL OFV.CERRARCONSULTA(RS)
			SQL ="SELECT * FROM PEDIDOS WHERE OBJID=0" & idpedido
				  ' RESPONSE.WRITE  "<BR>"  & SQL & "<BR>"
	        set RSDEC = ofv.crearconsultaEx(sql,cn,1,parametros)


			sql=" insert into PEDIDOS (SOLICITANTE,ESTADO,FECHAINI,CODOBJETOL,TIPOBJETOL,ACCION,RESPONSABLE,tipobjeto,objun) values ('"& SESSION.SESSIONID &"',1," & DHOY & "," & RSF("CODOBJETO") & ",'" & RSF("TIPOBJETO")&"','MODI','" & RESPONSABLE &"','" & TOBJ &"','" & SESSION("CLIEMP") & "')"
			call ofv.crearconsultaEx(sql,cn,1,parametros)
			Strsql = "Select * From PEDIDOS WHERE SOLICITANTE='" & SESSION.SESSIONID & "'"
			set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			IDPEDIDO=RS(0)
			CALL OFV.CERRARCONSULTA(RS)







            strsql="select * from objhtm where codobjeto=" & RSF("CODOBJETO")
        '    RESPONSE.WRITE STRSQL
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

            do while not rs.eof
             sql="insert into pedidosarch (idpedido,pos,descripcion,archivos) values("& IDPEDIDO & "," & rs("secuen") & ",'" & rs("archivo") & "','')"
             rs.movenext
            loop



            CALL OFV.CERRARCONSULTA(RS)



			sql=" UPDATE  PEDIDOS SET SOLICITANTE='" & SESSION("USUARIO") & "' WHERE OBJID="  & IDPEDIDO
' response.write sql & "<br>"
			call ofv.crearconsultaEx(sql,cn,1,parametros)
			COLS="IDPEDIDO"
			DELIM=","
			FOR I=3 TO 14
				IF NOT ISNULL(RSF(I)) THEN
					COLS=COLS & DELIM & RSF(I).NAME
				END IF
			NEXT

			COLS=COLS &",NOMBRE,PRODUCTO"
			VALORES= "'" & IDPEDIDO & "'"
			DELIM=","
			FOR I=3 TO 14
				IF NOT ISNULL(RSF(I)) THEN
					VALORES=VALORES & DELIM &  "'" & RSF(I) & "'"
				END IF
			NEXT

			VALORES=VALORES & ",'" & RSF(31) & "','" & RSF("TSEGMENTO") & "'"
			sql=" insert into PEDIDOSDET (" & COLS & ") values ("& VALORES & ")"
' response.write sql & "<br>"
			call ofv.crearconsultaEx(sql,cn,1,parametros)
			SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
					 set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
					 DESTINO=""
			IF NOT RS.EOF THEN DESTINO=SPLIT(RS("TEXTOIMPRESION"),"-")

			IF ISARRAY(DESTINO) THEN
				IF  ubound(DESTINO)=1 THEN
					SQL="INSERT INTO RELAOBJ(CODOBJMASTER,TIPOBJMASTER,CODOJASOC,TIPOBJASOC) VALUES (" & DESTINO(1) & ",'" &DESTINO(0) & "'," & IDPEDIDO & ",'"& RS(0)& "')"
				END IF
			END IF

			CALL OFV.CERRARCONSULTA(RS)
			CALL OFV.CERRARCONSULTA(RSDEC)
			call ofv.crearconsultaEx(sql,cn,1,parametros)
		END IF



	CASE "BAJA"

		SQL ="SELECT * FROM OBJPROPIEDADES  A "
		SQL= SQL & " INNER JOIN objprty B1 ON(A.CODOBJETO=B1.CODOBJETO and b1.estado=3) "
		SQL= SQL & " INNER JOIN objhtm B2 ON(A.CODOBJETO=B2.CODOBJETO ) "
		SQL= SQL & " INNER JOIN OBJETOS B ON(A.CODOBJETO=B.CODOBJETO)  WHERE  A.CODIGOFORM='" & SCODFORM & "'"
 '  RESPONSE.WRITE SQL
		SET RSf= ofv.crearconsultaEx(sql,cn,1,parametros)
        existe=false
		IF NOT RSf.EOF THEN
		    existe=true
			SQL="SELECT * FROM TIPOBJETO WHERE TIPOCPT='PD'"
			set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
			RESPONSABLE=""
			DESTINO=""
			IF NOT RS.EOF THEN
				RESPONSABLE=RS("WKFAPRO")
				DESTINO=SPLIT(RS("TEXTOIMPRESION"),"-")
				tobj=rs(0)
                CODOBJETO=RSF("CODOBJETO")
			end if

			CALL OFV.CERRARCONSULTA(RS)
			sql=" insert into PEDIDOS (SOLICITANTE,ESTADO,FECHAINI,CODOBJETOL,TIPOBJETOL,ACCION,RESPONSABLE,TIPOBJETO,objun) values ('"& SESSION.SESSIONID &"',1," & DHOY & "," & RSF("CODOBJETO") & ",'" & RSF("TIPOBJETO")&"','ELIM','" & RESPONSABLE &"','"& tobj &"','" & SESSION("CLIEMP") & "')"
'	 response.write sql & "<br>"

			call ofv.crearconsultaEx(sql,cn,1,parametros)
			Strsql = "Select * From PEDIDOS WHERE SOLICITANTE='" & SESSION.SESSIONID & "'"
			set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			IDPEDIDO=RS(0)
			CALL OFV.CERRARCONSULTA(RS)
			sql=" UPDATE  PEDIDOS SET SOLICITANTE='" & SESSION("USUARIO") & "' WHERE OBJID="  & IDPEDIDO
' response.write sql & "<br>"
			call ofv.crearconsultaEx(sql,cn,1,parametros)
			COLS="IDPEDIDO"
			DELIM=","
			FOR I=3 TO 14
				IF NOT ISNULL(RSF(I)) THEN
					COLS=COLS & DELIM & RSF(I).NAME
				END IF
			NEXT

			COLS=COLS &",NOMBRE,PRODUCTO"
			VALORES= "'" & IDPEDIDO & "'"
			DELIM=","
			FOR I=3 TO 14
				IF NOT ISNULL(RSF(I)) THEN
					VALORES=VALORES & DELIM &  "'" & RSF(I) & "'"
				END IF
			NEXT

			VALORES=VALORES & ",'" & RSF("DESCRI") & "','" & RSF("TSEGMENTO") & "'"
			sql=" insert into PEDIDOSDET (" & COLS & ") values ("& VALORES & ")"
' response.write sql & "<br>"
            CALL OFV.CERRARCONSULTA(RSf)
			call ofv.crearconsultaEx(sql,cn,1,parametros)


			IF ISARRAY(DESTINO) THEN
				IF  ubound(DESTINO)=1 THEN
					SQL="INSERT INTO RELAOBJ(CODOBJMASTER,TIPOBJMASTER,CODOJASOC,TIPOBJASOC) VALUES (" & DESTINO(1) & ",'" & DESTINO(0) & "'," & IDPEDIDO & ",'"& tobj & "')"
					call ofv.crearconsultaEx(sql,cn,1,parametros)
				END IF
			END IF


		END IF

	CASE "VER"
	    XPC=REQUEST.QUERYSTRING("XPC")
	   IF XPC &"" <> "" THEN
	      IDPEDIDO=XPC
	    ELSE
		IDPEDIDO= session("ncliente")
		END IF

		existe=true
		'response.write sdocmodal
		if sdocmodal="WKPED" THEN


		tabla1 =  "<div id='DTTEMAINT' name='DTTEMAINT' style='overflow:auto;"
        tabla1 = TABLA1 & "border-style:solid;border-width:1;border-color:#dddddd;' ></div>" & Chr(13)

		      WSPFRMN = "DTTEMAIT"
		        'WSPFRMA = "consdocumindex.asp"
		      '  WSPFRMA = "CONSTEMAIND.asp?FORMATO=ABR"
                 WSPFRMA="cowkfselect.asp?acc=ABR&SKIND=1&Tipobjeto=" & session("tcliente") &"&Codobjeto=" & session("ncliente")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
		         WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET=WK4 style='display:none;' METHOD='POST' ></FORM>"
'response.write session("ncliente") & "  " &  session("tcliente")


		        WSPDTTSCP = ""
		        WSPDTTSCP = WSPDTTSCP & "<script   >"
		       ' WSPDTTSCP = WSPDTTSCP &  " alert('ENTRO: ');"
		        '  WSPDTTSCP = WSPDTTSCP &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
		        WSPDTTSCP = WSPDTTSCP & "document." & WSPFRMN & ".submit();"
		        '''  WSPDTTSCP = WSPDTTSCP &  "document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").style.display="  & chr(34) & "block" & chr(34) & ";"
                WSPDTTSCP = WSPDTTSCP & "</script>"
		END IF
	END SELECT



	strsql = " select b.* ,a.accion,A.CODOBJETOL,A.TIPOBJETOL,a.objid as objidO,a.tipobjeto as tipobjetoO,a.estado as estadoP,A.SOLICITANTE from pedidos a"
	strsql=strsql & " inner join pedidosdet b on (a.objid=b.idpedido)"
	strsql= strsql & "   where b.idPEDIDO=0" & IDPEDIDO

'response.write strsql & "<BR>"

	set rsprop=ofv.crearconsultaEx(StrSql,cn,1,parametros)




	sImagenEM  = "<FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>&#0252;</FONT> "
	sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>4</FONT></B> "

	sHTML = sHTML &  "<div id='mt2' >"
	sForm = "edPropiedades"


	if not rsprop.eof then
        existe=true
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

		h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","","Modo  de Impresion: ","si")
		sdato = "---"


		SCOMSEP = chr(9)
		scomK = "SF" & SCOMSEP & "DF"
		scomD = "SIMPLE FAZ" & SCOMSEP & "DOBLE FAZ"

		sInput = ofv.gencomboFX(sForm,"MODOIMPRESION",rsprop("MODOIMPRESION"),scomK,scomD,"")

				 h1 =  h1 &  ofv.GenCelda("","","","45%","","","",sInput,"si")

				 h1 =  h1 &  ofv.GenCelda("",CLASEFILA,"","5%","","","","Firmas ","si")
				  sdato="---"
					  Strsql = "Select * From FORMADMTFIRMA where estado=3 Order By descriPCION"
		sdato = ofv.GenerarCombo(sForm ,rsPROP(9).name,rsPROP(9).value,Strsql,cn,0,1,"")



				 h1 =  h1 &  ofv.GenCelda("","","","45%","","","",sdato,"si")


		FNSHTMLP =  FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

				   sdato = "---"


		h1 =  ofv.GenCelda("",CLASEFILA,"","5%","","","","Tiempo&nbsp;Guarda:","si")

		sdato = "---"

		Strsql = "Select * From FORMADMTGUARDA where estado=3  Order By descriPCION"
		sdato = ofv.GenerarCombo(sForm ,rsPROP(4).name,rsPROP(4).value,Strsql,cn,0,1,"")

		h1 =h1 & ofv.GenCelda("","" ,"","40%","","","", sdato ,"si")


		h1 = h1 & ofv.GenCelda("",CLASEFILA,"","5%","","","","Producto:","si")
		Strsql = "Select * From TIPOPRODUCTO Order By descri"
		sproducto = ofv.GenerarCombo(SFORM,rsPROP(14).name,rsPROP(14).value,Strsql,cn,0,1,"")
		h1 = h1 & ofv.GenCelda("","","","45%","","","",sproducto,"si")
			         FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

	   sdato = "---"
	   h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Categoria","si")
	   sdato = "---"

        Strsql = "Select * From FORMADMCATG  where estado=3 Order By descriPCION"
		sdato = ofv.GenerarCombo(sForm ,rsPROP(6).name,rsPROP(6).value,Strsql,cn,0,1,"")

		h1 = h1 & ofv.GenCelda("","","","45%","","","",sdato,"si")


		sdato = "---"
		  h1 = h1 & ofv.GenCelda("",CLASEFILA,"","5%","","","", "Proceso","si")
		  sdato = "---"

		  Strsql = "Select * From FORMADMPROCESO where estado=3 Order By descriPCION"
		  SDATO = ofv.GenerarCombo(sForm ,rsPROP(12).name,rsPROP(12).value,Strsql,cn,0,2,"")

		h1 = h1 & ofv.GenCelda("","","","45%","","","",sdato,"si")

	  FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")




       sdato = "---"
	   h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Publicación en Conectados","si")

       SCOMSEP = chr(9)
	   		scomK = "SI" & SCOMSEP & "NO"
	   		scomD = "SI" & SCOMSEP & "NO"

		sInput = ofv.gencomboFX(sForm,"publiconectado",rsprop("publiconectado"),scomK,scomD,"")

        h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput ,"si")

		sdato = "---"
		  h1 = h1 & ofv.GenCelda("",CLASEFILA,"","5%","","","", "Motivo","si")
		  sdato = "---"

		  Strsql = "Select * From FORMADMMOTIVO where estado=3 Order By descriPCION"
		  SDATO = ofv.GenerarCombo(sForm ,"motivo",rsPROP("motivo"),Strsql,cn,0,2,"")

		h1 = h1 & ofv.GenCelda("","","","45%","","","",sdato,"si")

	    FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")


	    sdato = "---"
	   	   h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Asociado a proyecto","si")

	          SCOMSEP = chr(9)
	   	   		scomK = "SI" & SCOMSEP & "NO"
	   	   		scomD = "SI" & SCOMSEP & "NO"

	   		sInput = ofv.gencomboFX(sForm,"proyecto",rsprop("proyecto"),scomK,scomD,"")

	   		sInputD=""
		if rsprop("proyecto")="SI" THEN
		  sInputD = "<BR><textarea class=dt name=DESCRIPROYECTO id=DESCRIPROYECTO rows=4 cols=110 "
				if bModificar  then
					sInputD = sInputD & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
					sInputD = sInputD & " extenso'); return false;} else {"
					sInputD = sInputD & " document." & sform & ".submit() ;}" & chr(34)
				else
					sInputD = sInputD & " readonly "
				end if
		sInputD = sInputD & " >" &  trim(rsPROP("DESCRIPROYECTO")&"")  &"</textarea></font>"

        END IF


	           h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput & sInputD,"si")

             sdato = "---"
			 		  h1 = h1 & ofv.GenCelda("",CLASEFILA,"","5%","","","", "Tipo de Formulario","si")
			 		  sdato = "---"

			 		  Strsql = "Select * From FORMADMTFORMU where estado=3 Order By descriPCION"
			 		  SDATO = ofv.GenerarCombo(sForm ,"CLASEFORM",rsPROP("CLASEFORM"),Strsql,cn,0,2,"")

		h1 = h1 & ofv.GenCelda("","","","45%","","","",sdato,"si")


	      FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")


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
		'	FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

	        h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "DESCRIPCION","si")
		    sdato = "---"

	  sInput = "<textarea class=dt name=DESCRIPCION id=DESCRIPCION rows=4 cols=110 "
		if bModificar  then
			sInput = sInput & " onblur=" & chr(34) & " if(this.value.length==0) {this.value=' '}; if (this.value.length > 2300) {alert('El texto es demasiado"
			sInput = sInput & " extenso'); return false;} else {"
			sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
		else
			sInput = sInput & " readonly "
		end if

		xdsc=trim(rsPROP("DESCRIPCION") & "")

		if len(xdsc)=0 or xdsc="" then xdsc= " "

		sInput = sInput & " >" &  REPLACE(xdsc & "","'"," ")  &"</textarea></font>"
		h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput,"si")
		FNSHTMLP2 =  ofv.GenRow("","","","","","","",h1,"si")
		FNSHTMLP2= ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23; ' ","","98%","","0","0","0","0","","",FNSHTMLP2,"si")
        sacc2=sacc
		if sacc="VER" OR sacc="VERW" THEN
         xxx=getinfouser(rsPROP("SOLICITANTE"))
         sacc2=rsPROP("ACCION")
        ELSE
         xxx=getinfouser(session("usuario"))
        END IF

        h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>Dueño del formulario:&nbsp;" & session("SGU_nombre") &"</SPAN>","si")
		'sInput = ofv.GenerarInput(rsPROP(15).name,rsPROP(15),"text",SFORM,"100","dt","")
		'h1 = h1 & ofv.GenCelda("","","","45%","","","",sInput,"si")sacc

		FNSHTMLP0 =  ofv.GenRow("","xx bgcolor='"& Session("CMFMENU") &"'","","","45","","",h1,"si")

		 select case sacc2

		 case "ALTA"
		 h1 = ofv.GenCelda("","vt STYLE='text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'> ALTA DE FORMULARIO</SPAN>" ,"si")
		 CASE "MODI"
		 h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>MODIFICACION DE FORMULARIO</SPAN>" ,"si")
		 CASE "BAJA"
		  h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>BAJA DE FORMULARIO</SPAN>" ,"si")
		 CASE "ELIM"
		  h1 = ofv.GenCelda("","vt style='font-size:20px;text-align: center !important'","","98% colspan=2","","","", "<SPAN CLASS=ntslp style='font-size:20px;'>BAJA DE FORMULARIO</SPAN>" ,"si")
		 end select


		FNSHTMLP0 =  FNSHTMLP0 & ofv.GenRow("","xxxx bgcolor='"& Session("CMFMENU") & "'","","","45","","",h1,"si")

		 h1x = ofv.GenCelda("","","","98% colspan=2","","","", "&nbsp;" ,"si")
        FNSHTMLP0 =  FNSHTMLP0 & ofv.GenRow("","","","","","","",h1x,"si")
		h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "NOMBRE","si")
		sInput = ofv.GenerarInput(rsPROP(15).name,rsPROP(15),"text",SFORM,"100","dt","")
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
		set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		FNSHTMLA =  ofv.GenTabla("","aaa  style='border-style:solid;border-width:1;border-color:#f05f23;' ","","98%","",00,0,0,5,"","","","no")
		Celda = ofv.GenCelda("",clasefila & "  COLSPAN=2 ALIGN=LEFT","","","","","","<font size=2>ARCHIVOS/IMAGENES DEL FORMULARIO</font>","si")
        if Session("EMPMUN")  = "S" then
         sHref="uploadn.asp?acc=" &SACC2&"&id="&IDPEDIDO&"&direct=Emp"& SESSION("CLIEMP") &"\pedidos&nomdir=emp" & SESSION("CLIEMP") & "\pedidos" & "&seg=" & seg & "&vuel=PED"
        else
		 sHref="uploadn.asp?acc=" &SACC2&"&id="&IDPEDIDO&"&direct=pedidos&nomdir=pedidos" & "&seg=" & seg & "&vuel=PED"
		end if
		sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & SHREF & "'); return false;" & Chr(34)
		IF sACC <>  "VER" AND (SACC="ALTA" OR SACC="MODI") THEN
			LINKBSQ = ofv.GenLink("", "istt STYLE='CURSOR: hand;border-style:solid;border-SIZE:1;' ", sHref, "", "", sFuncion, "&nbsp;<B>SUBIR</B>&nbsp;", "si")
		ELSE
			LINKBSQ="&nbsp;"
		END IF

		Celda = CELDA & ofv.GenCelda("",clasefila & "  ALIGN=RIGHT ","","","","","",LINKBSQ,"si")

		FNSHTMLA = FNSHTMLA & ofv.GenRow("","","","","","","",Celda,"si")
		POS=1
		DO WHILE NOT rs.EOF

			If xTrColor = "#FFFFFF" Then
							xTrColor = "#FFFFFF" '"#F0F0F0"
					 Else
							xTrColor = "#FFFFFF"
			End If

			if bSuprimir and (  sACC <>  "VER") then
				sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
				sHref = sHref & "&Tabla=PEDIDOSARCH&volver=" & vuelta & "_pedido=" &sACC & "_idpedido=" & rsprop(1) & "_seg=" & seg & "&Eliminar=NSNC"
				Celda = ofv.BotonEliminar(rs(1),sHref)
				celda= replace(celda,"class="," title='Eliminar' class=")
			else
				Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
			end if


         '  IF SACC ="BAJA"  THEN
         '       stexto1 = "<A class=ttt  target=_new href='" & RS("DIRECTORIO") & "/"  & RS("ARCHIVO")& "'"
         '       stexto1 = stexto1 & "   >" & REPLACE(RS("DESCRI") &"",CHR(13),"<BR>") & "</a>"
		'	ELSE
		        if Session("EMPMUN")  = "S" then
		         stexto1 = "<A class=ttt  target=_new href='emp" & rs("objun") & "/" & rs("archivo")&"'"
		        else
			     stexto1 = "<A class=ttt  target=_new href='" &rs("archivo")&"'"
			    end if
			    stexto1 = stexto1 & "   >" & REPLACE(RS("DESCRIPCION") &"",CHR(13),"<BR>") & "</a>"

			'END IF


					 Celda = CELDA & ofv.GenCelda("",CLASEFILAT,"","90%","","","", stexto1,"si")




			FNSHTMLA = FNSHTMLA & ofv.GenRow("","ut  bgcolor='" & xTrColor & "'  ","","","","","",Celda,"si")

			RS.MOVENEXT
			POS=POS +1
		LOOP

		FNSHTMLA =FNSHTMLA & "</table>"
       ' RESPONSE.WRITE sACC
		if bModificar  AND SACC <> "BAJA" then
		    sAccion = "asp/ActualizarEN.asp?ID=" & rsprop(0) & "&Tabla=pedidosdet&volver=" & vuelta  & "_idpedido=" & rsprop(1) & "_seg=" & seg   & "_PEDIDO=" & SACC &"_docmodal=" & sdocmodal
			sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
			Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",FNSHTMLP,"si")
		ELSE
		Formulario =FNSHTMLP
		end if




		sHTML = sHTML & Formulario

		sHTML = sHTML  & FNSHTMLA


		IF sACC <>  "VER" AND sACC <>  "VERW"   THEN

			sForm="ALTA"

			if validardatos(rsprop,SACC) then
                           BTaccadd = " onClick=" & chr(34) &  "this.style.display='none';document.getElementById('cn').style.display='none'; document.getElementById('info').style.display='';document." & sForm & ".submit();" & chr(34) &  " "
                        else
                           IF SACC="ALTA" THEN
                           BTaccadd = " onClick=" & chr(34) &  "alert('DEBE COMPLETAR LOS CAMPOS DE NOMBRE Y ARCHIVO DEL FORMULARIO PARA CURSAR EL PEDIDO.');" & chr(34) &  " "
                           ELSEIF SACC="MODI" THEN
                             BTaccadd = " onClick=" & chr(34) &  "alert('Debe agregar un ATTACH para cursar el pedido.');" & chr(34) &  " "
                           ELSEIF SACC="BAJA" THEN
                             BTaccadd = " onClick=" & chr(34) &  "alert('Debe agregar un ATTACH para cursar el pedido.');" & chr(34) &  " "
                           ELSE


                           BTaccadd = " onClick=" & chr(34) &  "alert('Debe completar toda la informacion para cursar el pedido.');" & chr(34) &  " "
                           END IF

                        end if




                             SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CONFIRMAR  ' CLASS='BT' " & BTaccadd & " >"


			      SACCION="CREARPEDIDO.asp?PEDIDO=OK&idpedido=" & rsprop(1) & "&TIPOBJETO=" & tobj
		        	Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")
					'CONFIRMAR=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",Formulario,"si")
		        	Celda = ofv.gencelda("","ut","RIGHT","50%","","","",Formulario,"si")

			sForm="BAJA"
			SINPUT = "<INPUT id=cn  name=cn TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"



			SACCION="CREARPEDIDO.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)
			Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")
			Celda = CELDA & ofv.gencelda("","ut","LEFT","50%","","","",Formulario,"si")
			Filax  =  ofv.genrow("","AAA  " ,"","","","","",Celda,"si")
			TABLA=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")
			if rsprop("estadop") ="4" OR  sDOCMODAL="WKPED" then TABLA=""

		ELSE
			sForm="FORMACC"

           ' se fija accion del pedido alta,modi elim

      select case   rsprop("accion") & ""


           case "ALTA"

             	SACCION="asp/AgregarEsquema.asp?acc=ALTAPED&idpedido=" & idpedido
				SACCION="Copiar2.asp?BASEF=PEDIDO&acc=ALTA&idpedido=" & idpedido

				SELECT CASE RSPROP("ESTADOp")
				CASE 3

				SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CREAR FORMULARIO  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
				CASE 1
				SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
                SACCION="CREARPEDIDO.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)

				END SELECT

                IF  sACC <>  "VERW" THEN

				Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")
				ELSE
				Formulario=""
				END IF

				Celda = ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")



				SQL= "select * from CKTVIEW "
				SQL= SQL & " where  "
                SQL=SQL  & "OBJTIPO='"& rsprop("TIPOBJETOo") &"' and objcod='" &rsprop("objido") & "' and PSCFIN='NO' "
               ' response.write sql & sACC
                set rs5 = ofv.crearconsultaEx(SQL,cn5,1,parametros)

                IF NOT RS5.EOF and rsprop("estadop") <> 3 THEN
                        session("SEGCKTPSC")=RS5("SEGCKTPSC")

                        saccion="cowkfselect.asp?OPCION=RCH&ACC=CANCELA&SKIND=1&Tipobjeto=" &rsprop("TIPOBJETOo") &"&Codobjeto=" & rsprop("objido")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"

				 SINPUT = "<INPUT TYPE='BUTTON' VALUE=' CANCELA SOLICITUD  ' CLASS='BT' onClick='document." & sForm & "WKF" & ".submit();'>"
                  IF  sACC <>  "VERW" THEN

				 Formulario = ofv.GenForm(sForm & "WKF","",sAccion,"","POST","",SINPUT,"si")
                 ELSE
                 Formulario =""
                 END IF

			 	 Celda = CELDA & ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")
                END IF
				FILAX = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
				TABLA=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")
				if rsprop("estadop") ="4" OR  sDOCMODAL="WKPED" then TABLA=""

		  CASE "MODI"


			 	'SACCION="Copiar2.asp?BASEF=PEDIDO&acc=ALTA&idpedido=" & idpedido
			 	SACCION="CREARPEDIDO.asp?PEDIDO=MODIF&idpedido=" & idpedido & "&CODOBJETO=" & rsprop("CODOBJETOL") & "&TIPOBJETO=" & rsprop("TIPOBJETOL")

                 if rsprop("estadop") = 3 then
			 	SINPUT = "<INPUT TYPE='BUTTON' VALUE='  MODIFICAR FORMULARIO  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
		 elseif rsprop("estadop") = 1 then
			 	'SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
			 	SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
                               SACCION="CREARPEDIDO.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)
                 else
			 	SINPUT=""
		 end if



			 	Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")

			 	Celda = ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")

                SQL= "select * from CKTVIEW "
				SQL= SQL & " where  "
                SQL=SQL  & "OBJTIPO='"& rsprop("TIPOBJETOo") &"' and objcod='" &rsprop("objido") & "' and PSCFIN='NO' "
              '  response.write sql
                set rs5 = ofv.crearconsultaEx(SQL,cn5,1,parametros)

                IF NOT RS5.EOF and rsprop("estadop") <> 3 THEN
                   session("SEGCKTPSC")=RS5("SEGCKTPSC")

                   saccion="cowkfselect.asp?OPCION=RCH&ACC=CANCELA&SKIND=1&Tipobjeto=" &rsprop("TIPOBJETOo") &"&Codobjeto=" & rsprop("objido")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"

				 SINPUT = "<INPUT TYPE='BUTTON' VALUE=' CANCELA SOLICITUD  ' CLASS='BT' onClick='document." & sForm & "WKF" & ".submit();'>"

				 Formulario = ofv.GenForm(sForm & "WKF","",sAccion,"","POST","",SINPUT,"si")

			 	 Celda = CELDA & ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")
                END IF


			 	FILAX = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
				TABLA=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")

				 if rsprop("estadop") ="4" OR  sDOCMODAL="WKPED" then TABLA=""

		CASE "ELIM"


			 	'SACCION="Copiar2.asp?BASEF=PEDIDO&acc=ALTA&idpedido=" & idpedido
			 	SACCION="CREARPEDIDO.asp?PEDIDO=BAJAF&idpedido=" & idpedido & "&CODOBJETO=" & rsprop("CODOBJETOL") & "&TIPOBJETO=" & rsprop("TIPOBJETOL")

                 if rsprop("estadop") = 3 then
  			 	 SINPUT = "<INPUT TYPE='BUTTON' VALUE='  BAJA DE  FORMULARIO  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"

                 elseif rsprop("estadop") = 1 then
				'SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
				SINPUT = "<INPUT TYPE='BUTTON' VALUE='  CANCELAR  ' CLASS='BT' onClick='document." & sForm & ".submit();'>"
				 SACCION="CREARPEDIDO.asp?PEDIDO=ELIM&idpedido=" & rsprop(1)
		 else
				SINPUT=""

                 end if

			 	Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","",SINPUT,"si")

			 	Celda = ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")

                SQL= "select * from CKTVIEW "
				SQL= SQL & " where  "
                SQL=SQL  & "OBJTIPO='"& rsprop("TIPOBJETOo") &"' and objcod='" &rsprop("objido") & "' and PSCFIN='NO' "
              '  response.write sql
                set rs5 = ofv.crearconsultaEx(SQL,cn5,1,parametros)

                IF NOT RS5.EOF and rsprop("estadop") <> 3 THEN
                   session("SEGCKTPSC")=RS5("SEGCKTPSC")

                   saccion="cowkfselect.asp?OPCION=RCH&ACC=CANCELA&SKIND=1&Tipobjeto=" &rsprop("TIPOBJETOo") &"&Codobjeto=" & rsprop("objido")& "&Menu=&Sector=&MAINAPL=EM&GLBSOLN=S0&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"

				 SINPUT = "<INPUT TYPE='BUTTON' VALUE=' CANCELA SOLICITUD  ' CLASS='BT' onClick='document." & sForm & "WKF" & ".submit();'>"

				 Formulario = ofv.GenForm(sForm & "WKF","",sAccion,"","POST","",SINPUT,"si")

			 	 Celda = CELDA & ofv.gencelda("","ut","CENTER","50%","","","",Formulario,"si")
                END IF
			 	FILAX = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
				TABLA=ofv.GenTabla("","aaa  ","","100% ","","0","","0","0","","",FILAX,"si")
                if rsprop("estadop") ="4" OR  sDOCMODAL="WKPED" then TABLA=""



	end select

     end IF

end if
		sHTML = sHTML & "<br><br> " & TABLA
		sHTML = sHTML & "</div>"

        sInput="<B>GENERANDO  WORKFLOW DE SOLICITUD , AGUARDE POR FAVOR</B>"
	    celda =  ofv.GenCelda("","tt style='text-align: center !important;border-style:solid;border-width:1;border-color:#dddddd;'","","","50%","","",sInput,"si")
	    fila=ofv.GenRow("","","","","","","",celda,"si")
        TAB =  ofv.GenTabla("","","","98% id=info name=info  style='display:none;'","","0","","0","0","","",fila,"si")






		sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","") & TAB



	SHTML=SHTML & cerrarventana  & tabla1 & WSPFORME & WSPDTTSCP

	if not existe then
	  shtml = shtml & "<TABLE><TR><TD CLASS=UT><FONT color=#993333 size=3 face=WEBDINGS>i</FONT> NO EXISTE FORMULARIO CON EL CODIGO " & SCODFORM & "</TD></TR>"
	   SHREF="PedidoForm.ASP?Ing=P&BASE=K&seg=CLIMN52412655104952828166469"
	   sform="<form name=xsf id=xsf method=post action='" & SHREF & "'></form>"
	   sFuncion=" onclick='document.xsf.submit();' "
	   LINKBSQ2="<tr><td><BR> <input type=button  class=bietx	 xSTYLE='width:90%;height:50px;color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' " & sFuncion & " value='volver' ></td></tr></table>"
	   shtml=shtml & LINKBSQ2 & sform
	end if


	if sdocmodal="ED" THEN
     xsFormulario = "<form name=nuloz id=nuloz target=fesquemas action=" & Session("ASP") & "  method=post>"
     shtml=SHTML & xsFormulario & "<SCRIPT> "
	 shtml=SHTML & " window.top.CorWin.addEventListener('beforeunload', function (e) { "
	 shtml=SHTML & "          document.getElementById('nuloz').submit();"
	 shtml=SHTML & "          });"
     shtml=SHTML & "</script>"

	end if

	if sdocmodal="WKPED" THEN
     xsFormulario = "<form name=nuloz id=nuloz action='novedmenu.asp?tab=view4' target=info method=post>"
	 shtml=SHTML & xsFormulario & "<SCRIPT> "
	 shtml=SHTML & " window.top.CorWin.addEventListener('beforeunload', function (e) { "
	 shtml=SHTML & "          document.getElementById('nuloz').submit();"
	 shtml=SHTML & "          });"
	 shtml=SHTML & "</script>"

    end if


    if sdocmodal="CL" THEN
	     '  xsFormulario = "<form name=nuloz id=nuloz action='novedmenu.asp' target=info method=post>"
		   shtml=SHTML &  "<SCRIPT> function handleWindowClose(e) { "
		   shtml=SHTML & "   e = window.event || e; "
		 '  shtml=SHTML & "       if ((e.clientX < 0) || (e.clientY < 0))"
		'   shtml=SHTML & "       {"

		   shtml=SHTML & "         e.preventDefault();e.stopImmediatePropagation();e.stopPropagation();"
		'   shtml=SHTML & "        }"
		   shtml=SHTML & "}"
	    '   shtml=SHTML & " window.onbeforeunload = handleWindowClose;</script>"
    end if

	sHTML = sHTML & "<script> var element=document.getElementById('MTH'); if (typeof(element) != 'undefined' && element != null) {element.style.display='none';}</script>" & GTCCALENDAR("") & "</body></html>"
	call ofv.cerrarconn(cn)
	call ofv.cerrarconn(cn5)


end if


response.write sHTML

function validardatos(rspropx,ACC)




'IF ACC="BAJA" THEN
' validardatos=TRUE
' EXIT FUNCTION
'END IF

validardatos=false


select case  ACC

case "BAJA"

'     SQL="SELECT OBJID FROM PEDIDOSARCH WHERE IDPEDIDO=" & rspropx(1)

'	 	set rsATT = ofv.crearconsultaEx(sql,cn,1,parametros)

	' 	IF NOT rsATT.EOF THEN
	 	 validardatos=TRUE
	 '	END IF
	 '	CALL OFV.CERRARCONSULTA(rsATT)


      EXIT FUNCTION


CASE "ALTA"




	SQL="SELECT OBJID FROM PEDIDOSARCH WHERE IDPEDIDO=" & rspropx(1)
	' RESPONSE.WRITE SQL
	set rsATT = ofv.crearconsultaEx(sql,cn,1,parametros)

	if  trim(rspropx("nombre")& "") ="" then
	  exit function
	end if

	IF NOT rsATT.EOF THEN
	 validardatos=TRUE
	END IF
	CALL OFV.CERRARCONSULTA(rsATT)
	EXIT FUNCTION

CASE "MODI"
validardatos=TRUE

 SQL="SELECT OBJID FROM PEDIDOSARCH WHERE IDPEDIDO=" & rspropx(1)

 set rsATT = ofv.crearconsultaEx(sql,cn,1,parametros)
 IF   rsATT.EOF THEN
     validardatos=FALSE
 END IF
 CALL OFV.CERRARCONSULTA(rsATT)

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

CASE ELSE


if  rspropx("banca")& "" ="" then
  exit function
end if



if  rspropx("tiempoguarda")& "" ="" then
  exit function
end if


if  rspropx("catform")& "" ="" then
  exit function
end if


if  rspropx("firmasform")& "" ="" then
  exit function
end if

if  rspropx("modoimpresion")& "" ="" then
  exit function
end if



if  rspropx("proceso")& "" ="" then
  exit function
end if

if  rspropx("producto")& "" ="" then
  exit function
end if

if  rspropx("nombre")& "" ="" then
  exit function
end if

if  rspropx("publiconectado")& "" ="" then
  exit function
end if

if  rspropx("motivo")& "" ="" then
  exit function
end if

if  rspropx("tipoformu")& "" ="" then
  exit function
end if
validardatos=true
END SELECT

end function


%>
