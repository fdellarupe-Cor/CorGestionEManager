<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok = ofv.CheckUsuario()

If ok <> "Y" Then
    sHTML = ok
Else
    response.expires = 0

    response.buffer = True
   '   response.write request.querystring & "<br>" & request.form
    ''' verifica si es del circuito metodo 2.8.2 o anterior
    vuelta = ofv.vueltaasp2("../") & vuelta
	vuelta2 = ofv.vueltaasp2("")
    stipo = request.QUERYSTRING("Tipobjeto")
    nCodigo = request.QUERYSTRING("Codobjeto")
    SACC=request.QUERYSTRING("ACC")
    btnAP=request.QUERYSTRING("btnapr")
 '  response.write request.QUERYSTRING
	session("CKTMODAL")="N"
    Dim IdDomE
    IdDomE=1
    'response.write request.querystring & "<br>" & request.form
	'if session("CodigoN") = "" then
	'	session("CodigoN") = nCodigo
	'else
	'	if session("CodigoN") <> nCodigo and nCodigo <> "" then
	'		response.write ( "entrro por el if" )
	'		session("CodigoN") = nCodigo
	'	end if
	'end if
	'if session("TipoS") = "" then
	'	session("TipoS") = stipo
	'else
	'	if session("TipoS") <> stipo and stipo <> "" then
	'		session("TipoS") = stipo
	'	end if
	'end if


	'response.write (chr(13) & session("NCliente") & "3333" & chr(13))

    if session("ModMain")="WKPED" then
      tipocpt="PD"
    else
      tipocpt=""
    end if

    Set cn2 = ofv.conectar(ofv.strconn2)
    Set cn = ofv.conectar(ofv.strconn5)
	Set cn3 = ofv.conectar(ofv.strconn2)
    Set cn4 = ofv.conectar(ofv.strconn6)

    Strsql = "SELECT * FROM OBJWFL  "
    Strsql = Strsql & " Where codobjeto=0" & nCodigo
    Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
    Set rs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)

    If Not rs.EOF Then
        ''  response.write "v " & rs(6) & "<br>"
        If rs(6) Then
            sMenu = request.QUERYSTRING("Menu")
            sSector = request.QUERYSTRING("Sector")
            response.redirect "edObjWfl.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & stipo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=ESPWFL"
        End If
    End If
    Call ofv.cerrarconsulta(rs)

    mstdsc = ""
    WKFAUTO = "WKF0001693621131313343163"
    WKFAPRO = "WKF0001693673726313343163"
    DESLEVL = 9
    SUPLEVL = 9


    If session("TCliente") = "RSK" Then
        Strsql = "Select * From ORRSKCONFIG where APLIC = '" & WKAPLIC & "' "
        Set rs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)

        If Not rs.EOF Then
            WKFAUTO = rs(5)
            WKFAPRO = rs(6)
            DESLEVL = 9
            SUPLEVL = 8
        End If


        mstdsc = ""
        Strsql = "SELECT MSGCHAT FROM OBJWFL  "
        Strsql = Strsql & " Where codobjeto=" & session("NCliente")
        Strsql = Strsql & " And tipobjeto = '" & session("TCliente") & "'"

        Set rs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)

        If Not rs.EOF Then
            mstdsc = rs(0)
        End If

        XSPROCESO = ""
        Strsql = "SELECT * FROM  ORRSKCIRCUITO "
        Strsql = Strsql & " Where OBJID = 0" & session("NCliente")
        Set xrs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
        If Not xrs.EOF Then
            If xrs("APLIC") = "OR" Then
                TIPOOBJ = "CIRCUITO AUTOEVALUACION"
                XSPROCESO = GetProceso(xrs(1))
            Else
                TIPOOBJ = "PLAN DE CONTINGENCIAS"

            End If

            Strsql = "Select * From objetos where codobjeto = 0" & xrs("CODOBJETO")
            Set xrs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
            If Not xrs.EOF Then
                TIPOOBJ2 = xrs(1)
                TIPOOBJ3 = xrs(3)
                IdDomE=xrs("OBJUN")
            End If

            If WKAPLIC = "OR" Then
                session("NOMREQ") = TIPOOBJ & ":&nbsp;" & XSPROCESO    '& "&nbsp;-&nbsp;" & TIPOOBJ3
            Else
                session("NOMREQ") = TIPOOBJ & ":&nbsp;" & TIPOOBJ2 & "&nbsp;-&nbsp;" & TIPOOBJ3
            End If

        End If
        Call ofv.cerrarconsulta(xrs)
    Else
        tipocpt=""

        Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
        Set rs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
        If Not rs.EOF Then
            WKFAUTO = rs(19)
            WKFAPRO = rs(20)
            DESLEVL = 9
            SUPLEVL = rs(18)
            tipocpt=rs("tipocpt")
        End If

        mstdsc = ""
        Strsql = "SELECT MSGCHAT FROM OBJWFL  "
        Strsql = Strsql & " Where codobjeto=" & session("NCliente")
        Strsql = Strsql & " And tipobjeto = '" & session("TCliente") & "'"
        Set rs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)

        If Not rs.EOF Then
            mstdsc = rs(0)
        End If

        Strsql = "Select * From OBJETOS where CODOBJETO = " & session("NCliente")
        Set rs = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
        If Not rs.EOF Then
            Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
            Set rsX1 = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
            If Not rsX1.EOF Then
                TIPOOBJ = rsX1(1)
            End If

            if tipocpt = "K" then

            codForm = ""
            sql = "select codigoform from objpropiedades where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "' "
			Set rx = ofv.crearconsultaEx(sql, cn2, 1, parametros)

			if not rx.eof then codForm = rx(0)

				session("NOMREQ") = TIPOOBJ & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(3)

            else

            	session("NOMREQ") = TIPOOBJ & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(3)

        	end if
        End If
    End If

    if tipocpt="PD" or tipocpt="SN" THEN

    SQL= "SELECT   'Solicitud nro:' + cast (a.objid as nvarchar)  + ': ' + BA.NOMBRE FROM PEDIDOS A "
				    SQL= SQL  & "  INNER JOIN PEDIDOSDET BA ON(BA.IDPEDIDO=A.OBJID)"
				    SQL= SQL  & "  INNER JOIN TIPOBJETO B ON(B.TIPOBJETO=A.TIPOBJETO)"
				    SQL= SQL  & " WHERE A.OBJID=0" & session("NCliente")
				    'RESPONSE.WRITE SQL
				    'TIPOBJETO WHERE TIPOBJETO='" & TIPOBJETO & "'"
		             set rsx1 = ofv.crearconsultaEx(SQL,cn2,1,parametros)
		            if not rsx1.eof then

		                SESSION("NOMREQ")=UCASE(rsx1(0) & "")

                   end if
    END IF
    'response.write " msg:" & mstdsc & "-" & SESSION("NOMREQ") & "<br><br><br>"

    Call ofv.cerrarconn(cn2)

    '''''' fin verifica



    cantregmostrarx = 12
    cantregmoverx = CDbl(request.QUERYSTRING("CantRegMover"))


   ' Set cn = ofv.conectar(ofv.strconn5)
   ' Set cn3 = ofv.conectar(ofv.strconn2)
   ' Set cn4 = ofv.conectar(ofv.strconn6)

    ofv.ObtenerAtributos session("Form"), ""
    bAgregar = ofv.AAgregar
    bModificar = ofv.AModificar
    bSuprimir = ofv.AEliminar
    bConsultar = ofv.AConsultar
  '     response.write session("SEGCKT")
    SOPCION = request.QUERYSTRING("opcion")
    ACT = request.QUERYSTRING("act")


    sopx = SOPCION
    if Sacc="ABR" THEN
    WSPDIVN = "DTTEMAINT"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
     sHTML = ""
     sHTML = sHTML & "<HTML>"
	 '    sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"
     sHTML = sHTML & "<BODY  onload=""window.top.getIFD('Principal').getElementById('DTTEMAINT').innerHTML=document.body.innerHTML;"" >"
    ELSE

    sHTML = ofv.FormHeader(Replace(session("FormN"), "_", " "))
    'onpagehide
    sHTML=replace(shtml,"<body","<BODY onpagehide=""disp();""")
     shtml = shtml & CorPnlFnc("N","")
    sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

    sHTML = sHTML &  "<div id='mt2'  style='display:none;height:90%;overflow:auto;" & styl & "' >"
   END IF

   ' Strsql = "SELECT cktdsc FROM SEGCIRCUITO where segckt='" & session("SEGCKT") & "' "
   Strsql = "SELECT * FROM CKTVIEWEx WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	        Set rsx2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
	        If Not rsx2.EOF Then
		'        WKFAUTO = "WKF0001693621131313343163"
   				' WKFAPRO = "WKF0001693673726313343163"

	             select case cstr(rsx2(0) & "")
	             case "WKF00016631121722600921928","WKF00016631121722600921928DD"
	              wkft="AUTO"
	             CASE  "WKF00016631121722600921928999"
	              WKFt="APRO"
	             CASE else
	             wkf=""
	             end select
	            xtitulo=rsx2(7)
	             if tipocpt="PD" or tipocpt="SN" THEN
	                 SESSION("NOMREQ")=  SESSION("NOMREQ") '& "<br><br>" & rsx2("observaciones")
	             end if
	        else
	            xtitulo="&nbsp;"
	        End If
    Call ofv.cerrarconsulta(rsx2)
    'celdas = ofv.GenCelda("", "TTT", "", "", "", "", "", "AUTORIZACION DE DOCUMENTOS", "si")
    celdas = ofv.GenCelda("", "TTT", "", "", "", "", "", xtitulo, "si")
    filas = ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")
    sHTML = sHTML & ofv.GenTabla("", "aaa  ", "", "100%", "", "0", "", "-2", "-1", "", "", filas, "si")

    response.write sHTML

    filas = ""

    sHTML = ""

    response.flush

    vuelta = ofv.vueltaasp("../", cantregmostrarx, cantregmoverx) & "_OPCION=" & SOPCION & "_ACT=S"

   ' Dim ATABLA(200, 6)
    Dim ATABLA(200, 7)

    Dim ATABSW(200, 3)

    sImagenO = "<B><FONT FACE=CorFlaticon COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImageni = "<B><FONT FACE=CorFlaticon COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagene = "<B><FONT FACE=CorFlaticon COLOR=#f7aa22 SIZE=3>&#0232;</FONT></B>"
    sImagenC = "<B><FONT FACE=CorFlaticon COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenp = "<B><FONT FACE=CorFlaticon COLOR=#336699 SIZE=3>2</FONT></B>"
    sImagenRU = "<B><FONT FACE=CorFlaticon COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    sImagenRG = "<B><FONT FACE=CorFlaticon COLOR=#00aa00 SIZE=3>&#0148;</FONT></B>"
    sImagenRS = "<B><FONT FACE=CorFlaticon COLOR=#0000aa SIZE=3>G</FONT></B>"
    sImagenFI = "<B><FONT FACE=CorFlaticon COLOR=#6699cc SIZE=3>5</FONT></B>"
    sImagenff = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>6</FONT></B>"
    sImagenSE = "<B><FONT FACE=CorFlaticon COLOR=#6699cc SIZE=3>G</FONT></B>"
    sImagenMA = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>F</FONT></B>"

  '  sImagenA = "<B><FONT FACE=CorFlaticon COLOR=#00aa00 SIZE=3>&#0252;</FONT></B>"
   ' sImagenR = "<B><FONT FACE=CorFlaticon COLOR=#cc0000 SIZE=3>&#0251;</FONT></B>"
    sImagenA  = "<FONT FACE=Corflaticon COLOR=#005500 SIZE=3>&#097;</FONT>"
	sImagenR  = "<FONT FACE=Corflaticon COLOR=#AA0000 SIZE=4>&#0114;</FONT>"

    Strsql = "SELECT * FROM CKTVIEWEx WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
    Set RSX = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
    If Not RSX.EOF Then

        session("WKFPSC") = RSX(9)
        session("SEGCKT") = RSX(5)
        session("WKF") = RSX(0)
        session("ORGUSR") = RSX(15)
        IdDomE			   =RSX("CODEMPRE")

        ''''''''''''''''''''
        Set cn2 = ofv.conectar(ofv.strconn0)


        ''SESSION("WKFETPU") = "U"
        ''SESSION("WKFEUSR") = "COREMANAGER"
        ''SESSION("WKFEMAI") = "@"

        sInputx = "COR@MANAGER.COM"

        If session("WKFETPU") = "D" Then            ' EMISOR DEFAULT (EL RESPONSABLE DEL PASO)

            xxxx = getInfoUserw(session("ORGUSR"))
            If session("SGU_mail") <> "" Then sInputx = session("SGU_mail")

            sInputx = sInputx & ""

        ElseIf session("WKFETPU") = "U" Then        ' EMISOR PREDETERMINADO (PARAMETRIZADO EL LA BASE DE USUARIOS)

            WKFEUSR = session("WKFEUSR")

            If WKFEUSR = "" Then WKFEUSR = "COREMANAGER"


            Strsql = "Select EMAIL From USUARIOS where ID = '" & WKFEUSR & "' "
            Set rsniv = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
            If Not rsniv.EOF Then sInputx = rsniv(0)
            Call ofv.cerrarconsulta(rsniv)

        Else                                        ' EMISOR FIJO (USUARIO Y MAIL FIJOS EN CORCONECT)

            sInputx = session("WKFEMAI")


        End If



        If sInputx = "" Then sInputx = "COR@MANAGER.COM"

        session("ORGMAIL") = sInputx

        ''''''''''''''''''''''''''''''''

        session("CKTMODAL") = ""
        Strsql = "SELECT CKTMODAL FROM SEGCIRCUITO WHERE SEGCKT = '" & session("SEGCKT") & "' "
        Set rsx2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
        If Not rsx2.EOF Then
            session("CKTMODAL") = rsx2(0)
        End If
        Call ofv.cerrarconsulta(rsx2)

        If session("CKTMODAL") = "" Then session("CKTMODAL") = "S"

    End If
    Call ofv.cerrarconsulta(RSX)

    ''''''''''''''''''''''''''''''''''''
    '' busca originante y destinatario

    ATABLA(0, 0) = 2
    ATABLA(0, 1) = 0



    session("WKFCLASE") = "M"

    Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
    Set rs = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
    If Not rs.EOF Then
        session("WKFCLASE") = rs(3)
    End If


    If session("WKFCLASE") = "E" Then


        Strsql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & session("SEGCKT") & "' ORDER BY ORDEN, ID"
        '''   response.write StrSql
        Set rsxP1 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
        If Not rsxP1.EOF Then

            XINDT = 0

            Do While Not rsxP1.EOF

                XINDT = XINDT + 1

                ATABLA(XINDT, 0) = rsxP1(3)
                ATABLA(XINDT, 1) = rsxP1(5)
                ATABLA(XINDT, 2) = rsxP1(4)
                ATABLA(XINDT, 3) = rsxP1(2)
                ATABLA(XINDT, 4) = rsxP1(7)
                ATABLA(XINDT, 5) = rsxP1(8)
                ATABLA(XINDT, 7) = rsxP1("CODEMPRE")

                rsxP1.movenext
            Loop

            ATABLA(0, 0) = XINDT
            ATABLA(0, 1) = XINDT - 1

        End If
        Call ofv.cerrarconsulta(rsxP1)


    Else

        Strsql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & session("SEGCKT") & "' AND MODO = 'O'"
        '''   response.write StrSql
        Set rsxP1 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
        If Not rsxP1.EOF Then

            ATABLA(1, 0) = "O"
            ATABLA(1, 1) = rsxP1(5)
            ATABLA(1, 2) = rsxP1(4)
            ATABLA(1, 3) = "NO"
            ATABLA(1, 4) = rsxP1(7)
            ATABLA(1, 5) = rsxP1(8)
            ATABLA(1, 7) = rsxP1("CODEMPRE")

        End If
        Call ofv.cerrarconsulta(rsxP1)
        if tipocpt="PD" or tipocpt="SN" THEN
            Strsql = "select RESPONSABLE from segpaso where SEGCKT = '" & session("SEGCKT")  & "'  AND okrch='" & session("SEGCKTPSC") &  "'"

            Set rsxP1 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)

            IF NOT rsxP1.EOF  THEN
             Strsql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & session("SEGCKT") & "' AND CODIGO = '" & rsxP1(0) & "'"
            else
             Strsql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & session("SEGCKT") & "' AND MODO = 'D'"
            end if

          Call ofv.cerrarconsulta(rsxP1)

        ELSE

        Strsql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & session("SEGCKT") & "' AND MODO = 'D'"
        END IF

     '   response.write  Strsql
        Set rsxP1 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
        If Not rsxP1.EOF Then

            ATABLA(2, 0) = "D"
            ATABLA(2, 1) = rsxP1(5)
            ATABLA(2, 2) = rsxP1(4)
            ATABLA(2, 3) = "NO"
            ATABLA(2, 4) = rsxP1(7)
            ATABLA(2, 5) = rsxP1(8)
            ATABLA(2, 7) = rsxP1("CODEMPRE")

        End If
        Call ofv.cerrarconsulta(rsxP1)
        ''''''''''''''''''''''''''''''''''''

    End If

    'response.write (Chr(13))
    'response.write (session("WKFPSC") & "   " & nCodigo)
    'response.write (Chr(13))
	'response.write (SOPCION & "   " & ACT)
	'response.write (Chr(13))

	'Seteo una variable para enviar el mail una sola vez cuando se actualiza el paso siguiente y hay para aprobar mas de un paso.
	envioMail = false
'response.write "ggg<br>"
    If SOPCION = "OK" Or SOPCION = "RCH" Then

		' Comienza el if principal de opcion "S"
        If ACT = "S" Then
            sobserv = request.Form("OBS")

            'retornar para paso
              IF tipocpt="PD" or tipocpt="SN" THEN

		          SESSION("NOMREQ") = SESSION("NOMREQ") & "<BR><BR>" &  replace(sobserv & "" ,chr(13) ,"<br>")

               END IF
            RSPASO = ""
            xpaso=0
            DO while RSPASO = ""

            Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & session("WKFPSC") & "' "



            Set rs = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
'response.write Strsql & "ccc<br>"
            If Not rs.EOF Then ' comienzo de A1
                 if rs("clase")="M" OR RS("TIPO")="R" THEN RSPASO = "SALIR"

				'XMODEL = RS(5)
                'XPDSC = RS(6) & ":" & RS(7)
				'response.write(chr(13) & Strsql & "  "  & XMODEL  & "  " & XPDSC&  chr(13))
                PSCTIPO = rs(14)

                If SOPCION = "OK" Then
                    XTIPO = "APR"
                    XPASCOD = rs(12)
                Else
                    XTIPO = "RCH"
                    XPASCOD = rs(13)
                End If

                  ' ACA QUE ONDA?
                SEGCKTPSC = GTCKTPSC() & rs(6)

                'response.write session("NCliente") & chr(13)
				'response.write session("TCliente") & chr(13)

				'response.write(PSCTIPO)
'response.write "vvv<br>"
                If PSCTIPO = "U" Or (WKFAUTO = "WKF00016631121722600921928" and SOPCION = "RCH") Then
                    WXX = GENACCION(rs(5), XTIPO)
'response.write "zzz<br>"
                    'Proximo paso que se debería de hacer
                    Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
                  '  response.write strsql & "----" & chr(13)
                    Set RS2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                    If Not RS2.EOF Then

						Set cn2 = ofv.conectar(ofv.strconn0)
					'		response.Write(chr(13) & nCodigo & chr(13))

							STRSQL9 = "SELECT * FROM OBJWFLDET WHERE CODOBJETO = " & session("NCliente") & " AND tipobjeto = '" & session("TCliente") & "' and comen1 = '" & RS2("RCODIGO") & "'"
						'	response.Write(chr(13) & STRSQL9 & chr(13))
							Set RSPASOS = ofv.crearconsultaEx(STRSQL9, cn2, 1, parametros)
							If Not RSPASOS.EOF Then
								IF RSPASOS("ESTADOPASOWKF") = "N" THEN
							'		response.write (Chr(13) & Chr(13))
							'		response.write (RSPASOS("ORDEN") & "    " & RSPASOS("ESTADOPASOWKF") & "    " & XPASCOD)
							'		response.write (Chr(13) & Chr(13))
									session("WKFPSC") = XPASCOD
								ELSE
									envioMail = true
									RSPASO = "pase por "
								'	response.write(RSPASO & "   " & SOPCION & "  " & PSCTIPO) & chr(13)
								END IF
							ELSE
								envioMail = true
								RSPASO = "pase por "
						'		response.write(RSPASO & "   " & SOPCION & "  " & PSCTIPO) & chr(13)
							End If
						Call ofv.cerrarconn(cn2)

                        XMODEL = RS2(5)
                        XPDSC = RS2(6) & ":" & RS2(7)
						'response.write "xx<br>"
                        If rs(3) = "E" Then
                        'response.write "1<br>"
                            XRTIPO = RS2(15)
                            XRCOD = RS2(16)
                            XRASG = "NO"

                            XRLVL = 9
                            XRSCD = "0"


                            If XRTIPO = "USR" Or XRTIPO = "RSP" Then
                                XRASG = XRCOD
                                'response.write "2<br>"
                                If XRTIPO = "RSP" Then
                                    XRTIPO = "USR"
                                    For YYY = 1 To ATABLA(0, 0)
                                        If ATABLA(YYY, 0) = XRCOD Then
                                            XRCOD = ATABLA(YYY, 1)
                                            XRASG = XRCOD
                                            XRLVL = ATABLA(YYY, 4)
                                            XRSCD = ATABLA(YYY, 5)
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else

                            	XRTIPO = ATABLA(1, 2)
								XRCOD = ATABLA(1, 1)
								XRASG = ATABLA(1, 1)
								XRLVL = ATABLA(1, 4)
                            	XRSCD = ATABLA(1, 5)

                            End If
                        Else
                            If SOPCION = "RCH" Then
                            'response.write "3<br>"
                                XRTIPO = ATABLA(1, 2)
                                XRCOD = ATABLA(1, 1)
                                XRASG = ATABLA(1, 1)
                                XRLVL = ATABLA(1, 4)
                                XRSCD = ATABLA(1, 5)
                                XRID = ATABLA(1, 6)
                            Else
                            'response.write "4<br>"
                                XRTIPO = ATABLA(2, 2)
                                XRCOD = ATABLA(2, 1)
                                XRASG = "NO"
                                If XRTIPO = "USR" Then
                                    XRASG = XRCOD
                                End If
                                XRLVL = ATABLA(2, 4)
                                XRSCD = ATABLA(2, 5)
                                XRID = ATABLA(2, 6)
                            End If
                        End If



                        'response.write "1----"
                        'response.write "SEGCKTPSC: " & SEGCKTPSC & "<br>" & "XMODEL: " & XMODEL & "<br>" & "XRTIPO: " & XRTIPO & "<br>" & "XRCOD: " & XRCOD & "<br>" & "XRASG: " & XRASG & "<br>" & "XPDSC: " & XPDSC & "<br>" & "XRLVL: " & XRLVL & "<br>" & "XRSCD: " & XRSCD & "<br>"
                        WXX = GENCKTPSC(SEGCKTPSC, "NO", XMODEL, "NO", XRTIPO, XRCOD, XRASG, XPDSC, "", XRLVL, XRSCD,XRID)

                        '''''''''''''''''''
                        ' agregar envio de mail
                        ' SI EL PASO PASA A SER NO NO SE ENVÍA EL MAIL, AGREGAR CONDICIÓN
						if envioMail then
							if session("WKF") = "WKF00016631121722600921928" then
								mstdsc = sobserv
							end if
							xxxx = GENCKTMAIL(XRTIPO, XRCOD, mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@", session("ORGMAIL"), XRLVL)
							envioMail = false
							'response.write(chr(13) & " ENVIO EL MAIL " & chr(13))
						end if
                        '''''''''''''''''''
					else
						envioMail = true
						RSPASO = "pase por "
                    End If
                    Call ofv.cerrarconsulta(RS2)
                    If session("CKTMODAL") = "N" Then
                        Strsql = "SELECT SEGCKTPSC FROM SEGPASO WHERE SEGCKT = '" & session("SEGCKT") & "'"
                        Strsql = Strsql & " AND MODEL =  '" & session("WKFPSC") & "' AND FIN = 'NO' ORDER BY ID "
                        Set rsx2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                        If Not rsx2.EOF Then
                            Do While Not rsx2.EOF
                                If rsx2(0) <> session("SEGCKTPSC") Then
                                    ''''''''''''''''''''''''''''' ACTUALIZA PASO Y CIRCUITO DEL WORKFLOW CORRIENTE ''''''''''''''''''''''''''''''''''
                                    Strsql = "UPDATE SEGPASO SET FIN = '" & GETDH() & "', OKRCH = '" & SEGCKTPSC
                                    Strsql = Strsql & "', observaciones = 'CANCELADO POR RECHAZO DEL CIRCUITO.'"
                                    Strsql = Strsql & " WHERE SEGCKTPSC = '" & rsx2(0) & "' "
                                    Call ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                                End If
                                rsx2.movenext
                            Loop
                        End If
                        Call ofv.cerrarconsulta(rsx2)
                    End If

                Else

                    Set cn2 = ofv.conectar(ofv.strconn0)

                    sql = "update segpaso set segsalida='" & SOPCION & "' where segckt='" & session("SEGCKT") & "' and segcktpsc='" & session("SEGCKTPSC") & "'"
                    call ofv.crearconsultaex(sql,cn,1,parametros)

                    SNIV = 0
                    'strsql = "Select SECTOR From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
                    ' set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                    ' if not rsniv.eof then sniv = rsniv(0)
                    ' call ofv.cerrarconsulta(rsniv)
                    xxxx = getInfoUserw(session("USUARIO"))
                    If IsNumeric(session("SGU_idcostcenter")) Then SNIV = session("SGU_idcostcenter")


                    itab = 0
                    itabs = 0
                    Strsql = "Select * From SEGCKTRESP WHERE SEGCKTPSC = '" & session("WKFPSC") & "' AND SEGCKT = '" & session("SEGCKT") & "' order by ORDEN "
                             '       response.write StrSql
                    Set RS3 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                    If Not RS3.EOF Then
                        Do Until RS3.EOF
                            itab = itab + 1
                            ATABLA(0, 0) = ATABLA(0, 0) + 1
                            ATABLA(0, 1) = ATABLA(0, 1) + 1
                            '''          response.write itab

                            ATABLA(itab + 2, 0) = "L"
                            ATABLA(itab + 2, 1) = RS3(5)
                            ATABLA(itab + 2, 2) = RS3(4)
                            XWTIPO = RS3(4)
                            XWCOD = RS3(5)
                            XWSEC = RS3(8)
                            XWID = RS3(0)
                            ATABLA(itab + 2, 6) = RS3(0)
                            ATABLA(itab + 2, 3) = "NO"
                            ATABLA(itab + 2, 4) = RS3(7)
                            ATABLA(itab + 2, 5) = RS3(8)

                            ''   if (cstr(sniv) = cstr(rs3(5))) or (cstr(sniv) = cstr(rs3(8))) then

                            Select Case XWTIPO

                                Case "USR"

                                    If ucase(XWCOD) = session("USUARIO") Then
                                        itabs = itab
                                    End If

                                Case "GRP"

                                    If XWCOD = session("SEGGRP") Then
                                        itabs = itab
                                    End If

                                Case "SEC"

                                    If CStr(XWCOD) = CStr(SNIV) Then
                                        itabs = itab
                                    End If

                            End Select

                            RS3.movenext
                        Loop
                    End If

                    xmodalsen = "S"
                    If session("CKTMODAL") = "N" Then
                        ''    session("WKFPSC")    = RSX(9)
                        ''   SESSION("SEGCKT")    = RSX(5)

                        Strsql = "SELECT SEGCKTPSC FROM SEGPASO WHERE SEGCKT = '" & session("SEGCKT") & "'"
                        Strsql = Strsql & " AND MODEL =  '" & session("WKFPSC") & "' AND FIN = 'NO' ORDER BY ID "
                        Set rsx2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                        If Not rsx2.EOF Then
                            Do While Not rsx2.EOF
                                If rsx2(0) <> session("SEGCKTPSC") Then
                                    xmodalsen = "N"
                                    Exit Do
                                End If
                                rsx2.movenext
                            Loop
                        End If
                        Call ofv.cerrarconsulta(rsx2)

                        If xmodalsen = "S" Then

                            itabs = itab

                        Else

                            XMODEL = rs(5)
                            XPDSC = rs(6) & ":" & rs(7)

                            XRTIPO = ATABLA(itabs + 3, 2)
                            XRCOD = ATABLA(itabs + 3, 1)
                            XRASG = "NO"
                            XRLVL = ATABLA(itabs + 3, 4)
                            XRSCD = ATABLA(itabs + 3, 5)
                            XRID = ATABLA(itabs + 3, 6)


                        End If
                    End If

                    If xmodalsen = "S" Then

                        If itabs = itab Then

                            WXX = GENACCION(rs(5), XTIPO)
                            Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
                            Set RS2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                            If Not RS2.EOF Then

                                XMODEL = RS2(5)
                                XPDSC = RS2(6) & ":" & RS2(7)


                                If RS2(6) <> 999 Then
                                    XRTIPO = ATABLA(1, 2)
                                    XRCOD = ATABLA(1, 1)
                                    XRASG = ATABLA(1, 1)
                                    XRLVL = ATABLA(1, 4)
                                    XRSCD = ATABLA(1, 5)
                                    XRID = ATABLA(1, 6)
                                Else
                                    XRTIPO = ATABLA(2, 2)
                                    XRCOD = ATABLA(2, 1)
                                    XRASG = "NO"
                                    If XRTIPO = "USR" Then
                                        XRASG = XRCOD
                                    End If
                                    XRLVL = ATABLA(2, 4)
                                    XRSCD = ATABLA(2, 5)
                                    XRID = ATABLA(2, 6)
                                End If
                            End If
                            Call ofv.cerrarconsulta(RS2)

                        ElseIf itabs < itab Then
                            If itabs <> 0 Then
                                XMODEL = rs(5)
                                XPDSC = rs(6) & ":" & rs(7)

                                XRTIPO = ATABLA(itabs + 3, 2)
                                XRCOD = ATABLA(itabs + 3, 1)
                                XRASG = "NO" '& XRCOD
                                If XRTIPO = "USR" Then
								    XRASG = XRCOD
                                End If
                                XRLVL = ATABLA(itabs + 3, 4)
                                XRSCD = ATABLA(itabs + 3, 5)
                                XRID = ATABLA(itabs + 3, 6)

                            Else
                                WXX = GENACCION(rs(5), XTIPO)

                                Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
                                Set RS2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                                If Not RS2.EOF Then

                                    XMODEL = RS2(5)
                                    XPDSC = RS2(6) & ":" & RS2(7)


                                    If RS2(6) <> 999 Then
                                        XRTIPO = ATABLA(1, 2)
                                        XRCOD = ATABLA(1, 1)
                                        XRASG = ATABLA(1, 1)
                                        XRLVL = ATABLA(1, 4)
                                        XRSCD = ATABLA(1, 5)
                                        XRID = ATABLA(1, 6)
                                    Else
                                        XRTIPO = ATABLA(2, 2)
                                        XRCOD = ATABLA(2, 1)
                                        XRASG = "NO"
                                        If XRTIPO = "USR" Then
                                            XRASG = XRCOD
                                        End If
                                        XRLVL = ATABLA(2, 4)
                                        XRSCD = ATABLA(2, 5)
                                        XRID = ATABLA(2, 6)
                                    End If


                                End If
                                Call ofv.cerrarconsulta(RS2)
                            End If
                        End If

                        ''' GENERAR EL proximo PASO OK
                        'response.write "genera el proximo paso"
						'IDRESP = 0


								'sql = "select ID from segcktresp where segckt='" & session("segckt") & "' and modo='L'"
						       ' sql = sql & " and tipo='" & XRTIPO & "' and codigo='" & XRASG & "' order by id desc"
						        'response.write "sql: " & sql & "<br>"
						        'set xr = ofv.crearconsultaEx(sql,cn,1,parametros)

									'if not xr.eof then
									'	IDRESP = xr(0)
									'end if
						'response.write "2"
						'response.write "SEGCKTPSC: " & SEGCKTPSC & "<br>" & "XMODEL: " & XMODEL & "<br>" & "XRTIPO: " & XRTIPO & "<br>" & "XRCOD: " & XRCOD & "<br>" & "XRASG: " & XRASG & "<br>" & "XPDSC: " & XPDSC & "<br>" & "XRLVL: " & XRLVL & "<br>" & "XRSCD: " & XRSCD & "<br>"
                        WXX = GENCKTPSC(SEGCKTPSC, "NO", XMODEL, "NO", XRTIPO, XRCOD, XRASG, XPDSC,"", XRLVL, XRSCD,XRID)

						sql = "select responsable from segcircuito where segckt='" & session("segckt") & "'"
						set rp = ofv.crearconsultaex(sql,cn,1,parametros)

						'strsql = "select Count(*) from segpaso where segckt='" & session("segckt") & "' and segsalida=''"
                       ' set rx = ofv.crearconsultaex(strsql,cn,1,parametros)

                        ' agregar envio de mail
						if session("WKF") = "WKF00016631121722600921928DD" then

                        	if (sopcion = "RCH" or sopcion = "OK") and (rp(0) <> XRCOD) then
                        	 	mstdsc = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un CIRCUITO"
             					mstdsc = mstdsc & " pendiente de su aprobación.<br><br>SOLICITUD DE AUTORIZACION" & " POR @@@NORMREQ@@@"
                        	 	'XPDSC = "SOLICITUD DE AUTORIZACION"
                        	else
                        		mstdsc = sobserv  & "<br><br>" & "RECHAZO DE RESPONSABLES" & " @@@NORMREQ@@@"
                        	end if

                        else

							if tipocpt="PD" or tipocpt="SN" then

								mstdsc = "20: AUTORIZACION DE RESPONSABLES POR @@@NORMREQ@@@"

							end if

                        end if
                       SELECT CASE  session("APLMDCLI")

                       CASE "BPN"
                          IF SOPCION = "RCH" or (SOPCION = "OK" and session("WKF") = "WKF00016631121722600921928PED") THEN
						  	 xxxx = GENCKTMAIL(XRTIPO, XRCOD, mstdsc, session("ORGMAIL"), XRLVL)
                          ELSE
                             xxxx = GENCKTMAIL(XRTIPO, XRCOD, mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@", session("ORGMAIL"), XRLVL)
                          END IF
                        CASE ELSE
						   if SOPCION = "RCH" or (SOPCION = "OK" and session("WKF") = "WKF00016631121722600921928PED") then
							 xxxx = GENCKTMAIL(XRTIPO, XRCOD, mstdsc, session("ORGMAIL"), XRLVL)
                           end if

                        END SELECT
					call ofv.cerrarconsulta(rp)
					call ofv.cerrarconsulta(rx)
                    End If

                End If

				IF WKFAUTO = "WKF00016631121722600921928" THEN
					sql = "update segpaso set segsalida='" & SOPCION & "' where segckt='" & session("SEGCKT") & "' and segcktpsc='" & session("SEGCKTPSC") & "'"
                    call ofv.crearconsultaex(sql,cn,1,parametros)
				END IF

                ''''''''''''''''''''''''''''' ACTUALIZA PASO Y CIRCUITO DEL WORKFLOW CORRIENTE ''''''''''''''''''''''''''''''''''
                wkmsg = "<OK>"
                if sopcion = "RCH" then
                   wkmsg = "<RCH>"
                   Strsql = "SELECT responsable FROM segcircuito WHERE SEGCKT = '" & session("SEGCKT") & "'"

				Set RSX = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
            	If Not RSX.EOF Then

                	if XRTIPO = "USR" and ( XRASG <> rsx(0) ) then
                        '''''''''''''''''''
					' agregar envio de mail
					'response.write "XRTIPO: " & XRTIPO & "<br>" & "session(ORGMAIL): " & session("ORGMAIL")
				 	 'response.write "ABAJO " & rsx(0)
				 	xxxx = GENCKTMAIL(XRTIPO, rsx(0), sobserv & "<br><br>" & "RECHAZO DE RESPONSABLES" & " @@@NORMREQ@@@", session("ORGMAIL"), XRLVL)
                        '''''''''''''''''''
					end if
				end if
                end if



                Strsql = "UPDATE SEGPASO SET FIN = '" & GETDH() & "', OKRCH = '" & SEGCKTPSC & "', observaciones = '" & sobserv & "' WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
                'response.write strsql
                Call ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                'response. write Strsql & "------------------------bnbnbn" & cbool(rs(11)) & chr(13)
                If rs(11) Then

                    Strsql = "UPDATE SEGCIRCUITO SET FIN = '" & GETDH() & "'  WHERE SEGCKT = '" & session("SEGCKT") & "' "
                    Call ofv.crearconsultaEx(Strsql, cn, 1, parametros)
					envioMail = true
					RSPASO = "pase por "
				else
					'cambiar session para que sea el próximo paso
					if RSPASO<>"SALIR" THEN

					 session("SEGCKTPSC") = SEGCKTPSC
					END IF
                End If
			'else
			'	envioMail = true
			'	RSPASO = "pase por "
            End If 'finaliza A1
           ' xpaso=xpaso+1
          '  response.write xpaso & chr(13) & "--" & RSPASO & "--" & chr(13)
			loop ' finaliza el loop para saber si se tiene que saltear y aprobar algun paso mas.

            Strsql = "SELECT * FROM CKTVIEWEx WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "




            Set RSX = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
            If Not RSX.EOF Then


                celdas = ofv.GenCelda("", "TTr colspan=2", "", "", "", "", "", "WORKFLOW", "si")
                filas = ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "CIRCUITO:", "si")
                celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenC & "</b> " & RSX(7), "si")
                filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "PASO:", "si")
                'if session("APLMDCLI") = "BGAL" and session("TCliente") = "SN" and session("WKF") = "WKF00016631121722600921928PEDSOL" and SOPCION = "RCH" then
				'	celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenp & "</b> 30: RECHAZO DE RESPONSABLES", "si")
                'else
                	celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenp & "</b> " & RSX(16), "si")
                'end if
                filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "OPCION:", "si")
                If SOPCION = "OK" Then
                    SOPCDESC = " <b>" & sImagenA & "</b> " & "APRUEBA"
                Else
                    SOPCDESC = " <b>" & sImagenR & "</b> " & "RECHAZA"
                End If

                celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", SOPCDESC, "si")
                filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                If RSX(11) <> "NO" Then

                    celdas = ofv.GenCelda("", "vt colspan=2 style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ", "CENTER", "", "", "", "", "<B>PASO FINALIZADO:</B>", "si")
                    celdas2 = ofv.GenCelda("", "vt colspan=2", "CENTER", "", "", "", "", "<b>" & sImagenff & "</b> " & RSX(11), "si")

                    filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")
                    filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas2, "si")


                End If



                sHTML = sHTML & ofv.GenTabla("", "aaa  ", "", "100%", "", "0", "", "-2", "-1", "", "", filas, "si") & "<BR><BR>"

            End If





		' Finaliza el if principal de opcion "S"
        Else


            Strsql = "SELECT * FROM CKTVIEWEx WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "

		'	response.write(strsql & chr(13) & "123")

            Set RSX = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
            If Not RSX.EOF Then


                celdas = ofv.GenCelda("", "TTr colspan=2", "", "", "", "", "", "WORKFLOW", "si")
                filas = ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "CIRCUITO:", "si")
                celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenC & "</b> " & RSX(7), "si")
                filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "PASO:", "si")
                'if session("APLMDCLI") = "BGAL" and session("TCliente") = "SN" and session("WKF") = "WKF00016631121722600921928PEDSOL" and SOPCION = "RCH" then
				'	celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenp & "</b> 30: RECHAZO DE RESPONSABLES", "si")
                'else
                	celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenp & "</b> " & RSX(16), "si")
                'end if
                filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "OPCION:", "si")
                If SOPCION = "OK" Then
                    SOPCDESC = " <b>" & sImagenA & "</b> " & "APRUEBA"
                Else
                    SOPCDESC = " <b>" & sImagenR & "</b> " & "RECHAZA"
                End If

                celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", SOPCDESC, "si")
                filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")


                sHTML = sHTML & ofv.GenTabla("", "aaa  ", "", "100%", "", "0", "", "-2", "-1", "", "", filas, "si") & "<BR><BR>"

            End If

			'rutina cambia titulo textarea..envio mail o observaciones
			puede = 1

			if tipocpt = "SN" then

				sql = "select codigo,orden from segcktresp where segckt='" & session("segckt") & "' and orden <> '999' order by id desc"
				set rz = ofv.crearconsultaEx(sql,cn,1,parametros)

				if not rz.eof then
					if session("USUARIO") = rz(0) then
						puede = 0
					end if
				end if

			end if

			if puede = 1 then
            	celda = ofv.GenCelda("", "vt colspan=3 style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;' ", "center", "100%", "", "", "", "OBSERVACIONES/COMENTARIOS", "si")
            else
				celda = ofv.GenCelda("", "vt colspan=3 style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;' ", "center", "100%", "", "", "", "TEXTO MAIL PARA LOS DESARROLLADORES", "si")
		    end if

            filas = ofv.GenRow("", "", "", "", "", "", "", celda, "si")

            nCol = 80

            wdescri = "EL AUTORIZANTE NO REGISTRA OBSERVACIONES"

            t3b = "<font class=at><textarea class=dt "
            t3b = t3b & "name=OBS cols=" & nCol & " rows=20 "
            If bModificar Then
                '''  t3b = t3b & " onblur=Usuarios.submit() "
                t3b = t3b & " onkeypress=" & Chr(34)
                t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
                t3b = t3b & " extenso'); return false;}" & Chr(34) & " >"
            Else
                t3b = t3b & " readonly >"
            End If
            xdatos = wdescri
            t3b = t3b & xdatos & "</textarea></font>"
            celda = ofv.GenCelda("", "abc colspan=3", "center", "100%", "", "", "", t3b, "si")
            filas = filas & ofv.GenRow("", "", "", "", "", "", "", celda, "si")

            sinput = "&nbsp;"
            celda = ofv.GenCelda("", "", "center", "", "50%", "", "", sinput, "si")

         '' Agregado 20042015
         '''sinput = "<input CLASS=BT type=submit value='CONTINUAR' size=10 >"
            sinput = "<input CLASS=BT type=submit value='CONTINUAR' size=10 onclick=""this.style.cursor='wait';this.style.visibility='hidden';""  >"
         '' Agregado 20042015 FIN

            XVUELTA = Replace(request.QUERYSTRING, "OPCION=OK&", "")
            XVUELTA = Replace(request.QUERYSTRING, "OPCION=OK", "")
            XVUELTA = Replace(XVUELTA, "OPCION=RCH&", "")
            XVUELTA = Replace(XVUELTA, "&acc=ABR", "acc=ABR")
           ' response.write  XVUELTA

            sFunc = "document.Usuarios.action='COWKFselect.ASP?" & XVUELTA & "';document.Usuarios.submit();"

            SINPUT2 = ofv.GenerarInput("", "VOLVER", "button", "", "10", "bt", sFunc)

            celda = celda & ofv.GenCelda("", "", "center", "", "1%", "", "", sinput & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & SINPUT2, "si")

            sinput = "&nbsp;"
            celda = celda & ofv.GenCelda("", "", "center", "", "1%", "", "", sinput, "si")
            filas = filas & ofv.GenRow("", "", "", "", "", "", "", celda, "si")

            filas=ofv.GenTabla("", "aaa  ", "", "100%", "", "0", "", "-2", "-1", "", "", filas, "si")
            If bModificar Then
                sAccion = "COWKFselect.ASP?ACC=" &sacc&"&CantRegMover=0&OPCION=" & SOPCION & "&ACT=S"
                TARGET=""
                IF SACC="ABR" THEN TARGET=" TARGET=WK4"
                formulario0 = ofv.GenForm("Usuarios" & TARGET, "", sAccion, "", "post", "", filas, "si") & "<BR><BR>"
            Else
                formulario0 = ofv.GenForm("", "", "", "", "post", "", filas, "si")
            End If

           '' formulario0 = ofv.GenTabla("", "aaa  ", "", "100%", "", "0", "", "-2", "-1", "", "", filas, "si")

            sHTML = sHTML & formulario0 & "<BR><BR>"
             'ofv.GenTabla("", "aaa  ", "", "100%", "", "0", "", "-2", "-1", "", "", formulario0, "si") & "<BR><BR>"

        End If

'RUTINA ENVIO DE MAIL A LOS DESARROLLADORES SI ES SOLICITUD Y ULTIMO PASO DEL WKF
	if tipocpt = "SN" and SOPCION = "OK" and ACT = "S" then

		tsql = "select codigo from segcktresp where segckt='" & session("segckt") & "' order by id desc"
		set tz = ofv.crearconsultaEx(tsql,cn,1,parametros)

		if not tz.eof then

			if tz("codigo") = session("USUARIO") then

				Set cn2 = ofv.conectar(ofv.strconn2)
				sSql = "select usr from objwfldet where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "'"
				Set em = ofv.crearconsultaEx(sSql, cn2, 1, parametros)

				do while not em.eof

					xEnv = GENCKTMAIL("USR",em("usr"),session("NOMREQ"),session("ORGMAIL"),"3")

				em.movenext
				loop

			end if

		end if

	end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Else 'If de OK o rechazado



        Strsql = "SELECT * FROM CKTVIEWEx WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
        		'response.write strsql & "<br>" & chr(13)
        Set RSX = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
        If Not RSX.EOF Then


            celdas = ofv.GenCelda("", "TTr colspan=2", "", "", "", "", "", "WORKFLOW", "si")
            filas = ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

            celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "CIRCUITO:", "si")
            celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenC & "</b> " & RSX(7), "si")
            filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

            celdas = ofv.GenCelda("", "vt", "", "15%", "", "", "", "PASO:", "si")
            celdas = celdas & ofv.GenCelda("", "vt", "", "80%", "", "", "", "<b>" & sImagenp & "</b> " & RSX(16), "si")
            filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

            sHTML = sHTML & ofv.GenTabla("", "AAA ", "", "98%", "", "0", "", "0", "0", "", "", filas, "si") & "<BR>"




            Strsql = "SELECT * FROM CKTVIEWEx WHERE OKRCH = '" & session("SEGCKTPSC") & "' "
            Set rsXT = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
            ''     response.write strsql & "<br>" & chr(13)
            If Not rsXT.EOF Then

                PSCMODEL = rsXT(9)
                PSCMOT = "A"

                Strsql = "SELECT RCHPSC FROM WKFPASOS WHERE WKFPASCOD = '" & PSCMODEL & "' "
                Set RSXU = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                If Not RSXU.EOF Then

                    If RSXU(0) = RSX(9) Then PSCMOT = "R"

                End If

                Select Case PSCMOT
                    Case "A"
                        PSCMOT = sImagenA & "&nbsp;POR APROBACION&nbsp;&nbsp;"
                    Case "R"
                        PSCMOT = sImagenR & "&nbsp;POR RECHAZO&nbsp;&nbsp;"
                End Select



                'celdas = ofv.GenCelda("", "VT ", "", "50%", "", "", "", "VIENE DE <BIG><B>" & rsXT(16) & "</B></BIG>", "si")
                'celdas = celdas & ofv.GenCelda("", "VT ", "RIGHT", "50%", "", "", "", PSCMOT & "<SMALL><SMALL>" & rsXT(11) & "</SMALL></SMALL>", "si")
                'filas = ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                Set cn2 = ofv.conectar(ofv.strconn0)

                SNIV = "DESCONOCIDO"
                SNSEC = "DESCONOCIDO"

                SELECT CASE session("APLMDCLI")


			    CASE "GP"

			     IF rsXT("codempre") & "" <>  "" THEN IdDomE= rsXT("codempre")

                xxxx = getInfoUserweX(rsXT(15),IdDomE)
			    CASE ELSE
                 xxxx = getInfoUserw(rsXT(15))

                END SELECT

                SNIV = session("SGU_nombre")
                If IsNumeric(session("SGU_idcostcenter")) Then
                    SNSEC = GETRESPOCKT(session("SGU_idcostcenter"), "SEC")
                End If


                'celdas = ofv.GenCelda("", "vt colspan=2", "", "", "", "", "", "RESPONSABLE:&nbsp;" & sImagenRU & "&nbsp;" & SNIV, "si")
                'filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")
                'celdas = ofv.GenCelda("", "vt colspan=2 ", "", "", "", "", "", "AREA/SECTOR:&nbsp;" & sImagenSE & "&nbsp;" & SNSEC, "si")
                'filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")


                'filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                'nCol = 110
                't3b = "<font class=at><textarea style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#f7aa22;background-color:#ffffff;color:#000080;' "
                't3b = t3b & "name=OBS cols=" & nCol & " rows=4 "
                't3b = t3b & " readonly >"
                't3b = t3b & rsXT(17) & "</textarea></font>"


                'celdas = ofv.GenCelda("", "vt COLSPAN=2  STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ", "center", "", "", "", "", t3b, "si")
                'filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas, "si")

                'sHTML = sHTML & ofv.GenTabla("", "AAA STYLE='BORDER-STYLE:ridge;BORDER-WIDTH:2;border-color:#f7aa22;' ", "", "98%", "", "0", "", "0", "0", "", "", filas, "si") & "<BR><BR>"


            End If


            session("WKFPSC") = RSX(9)
            session("SEGCKT") = RSX(5)

            If RSX(11) <> "NO" Then

                celdas1 = ofv.GenCelda("", "vt style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ", "CENTER", "", "", "", "", "<B>PASO FINALIZADO:</B>", "si")
                celdas2 = ofv.GenCelda("", "vt", "CENTER", "", "", "", "", "<b>" & sImagenff & "</b> " & RSX(11)   , "si")


                filas = ofv.GenRow("", "", "", "", "10", "", "", celdas1, "si")
                filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas2, "si")



                sHTML = sHTML & ofv.GenTabla("", "aaa  ", "", "100%", "", "0", "", "-2", "-1", "", "", filas, "si") & "<BR><BR>"



            Else

                Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RSX(9) & "' "
                Set RS2 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                If Not RS2.EOF Then

                    PSCTIPO = RS2(14)
                    pscclase = RS2(3)

                    '''     response.write "paso n1<br>"


                    '''response.write WSORIG & " - " & WUORIG & " - " & WSDEST & " - " & WUDEST & "<br>"


                    If (RS2(8) Or RS2(9) Or RS2(11)) And RS2(12) <> "" Then

                        '''  response.write "paso x1<br>"

                        If PSCTIPO = "U" Then

                            Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(12) & "' "
                            Set RS3 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                            If Not RS3.EOF Then
                                sinput = RS3(6) & ":" & RS3(7)


                                If pscclase = "E" Then

                                    XRTIPO = RS3(15)
                                    XRCOD = RS3(16)
                                    XRASG = "NO"

                                    XRLVL = 9
                                    XRSCD = "0"

                                    If XRTIPO = "USR" Or XRTIPO = "RSP" Then
                                        XRASG = XRCOD
                                        If XRTIPO = "RSP" Then
                                            XRTIPO = "USR"
                                            For YYY = 1 To ATABLA(0, 0)

                                                If ATABLA(YYY, 0) = XRCOD Then
                                                    XRCOD = ATABLA(YYY, 1)
                                                    XRASG = XRCOD
                                                    XRLVL = ATABLA(YYY, 4)
                                                    XRSCD = ATABLA(YYY, 5)
                                                    Exit For
                                                End If

                                            Next

                                        End If
										Else

													XRTIPO = ATABLA(1, 2)
													XRCOD = ATABLA(1, 1)
													XRASG = ATABLA(1, 1)
													XRLVL = ATABLA(1, 4)
													XRSCD = ATABLA(1, 5)

										End If

									CKTTIPO = XRTIPO
									CKTCOD = XRCOD
									CKTUSR = XRASG

									SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
									SINPUT3 = GETRESPOCKT(CKTUSR, "USR")

                                Else

                                    XRTIPO = ATABLA(2, 2)
                                    XRCOD = ATABLA(2, 1)
                                    XRASG = "NO"
                                    If XRTIPO = "USR" Then
                                        XRASG = XRCOD
                                    End If
                                    XRLVL = ATABLA(2, 4)
                                    XRSCD = ATABLA(2, 5)


                                End If


                                CKTTIPO = XRTIPO
                                CKTCOD = XRCOD
                                CKTUSR = XRASG


                                SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
                                SINPUT3 = GETRESPOCKT(CKTUSR, "USR")

                            Else
                                sinput = "S/D"
                                SINPUT2 = "S/D"
                                SINPUT3 = "S/D"
                            End If
                        Else
                            Set cn2 = ofv.conectar(ofv.strconn0)

                            SNIV = 0

                            'strsql = "Select SECTOR From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
                            'set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                            'if not rsniv.eof then sniv = rsniv(0)
                            'call ofv.cerrarconsulta(rsniv)

                            xxxx = getInfoUserw(session("USUARIO"))

                            If IsNumeric(session("SGU_idcostcenter")) Then SNIV = session("SGU_idcostcenter")

                            itab = 0
                            itabs = 0
                            Strsql = "Select * From SEGCKTRESP WHERE SEGCKTPSC = '" & RSX(9) & "' AND SEGCKT = '" & session("SEGCKT") & "' order by ORDEN "
                            Set RS3 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                            If Not RS3.EOF Then
                                Do Until RS3.EOF
                                    itab = itab + 1

                                    ATABLA(0, 0) = ATABLA(0, 0) + 1
                                    ATABLA(0, 1) = ATABLA(0, 1) + 1
                                    '''          response.write itab

                                    ATABLA(itab + 2, 0) = "L"
                                    ATABLA(itab + 2, 1) = RS3(5)
                                    ATABLA(itab + 2, 2) = RS3(4)
                                    ATABLA(itab + 2, 3) = "NO"
                                    ATABLA(itab + 2, 4) = RS3(7)
                                    ATABLA(itab + 2, 5) = RS3(8)

                                    XWTIPO = RS3(4)
                                    XWCOD = RS3(5)
                                    XWSEC = RS3(8)

                                    ''   if (cstr(sniv) = cstr(rs3(5))) or (cstr(sniv) = cstr(rs3(8))) then
                                    ''      itabs = itab
                                    ''   end if

                                    Select Case XWTIPO

                                        Case "USR"


                                            If ucase(XWCOD) = session("USUARIO") Then
                                                itabs = itab
                                            End If


                                        Case "GRP"

                                            If XWCOD = session("SEGGRP") Then
                                                itabs = itab
                                            End If

                                        Case "SEC"

                                            If CStr(XWCOD) = CStr(SNIV) Then
                                                itabs = itab
                                            End If

                                    End Select





                                    RS3.movenext
                                Loop
                            End If
                            If itabs = itab Then
                                Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(12) & "' "
                                Set RS3 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                                If Not RS3.EOF Then
                                    sinput = RS3(6) & ":" & RS3(7)


                                    If pscclase = "E" Then

                                        XRTIPO = RS3(15)
                                        XRCOD = RS3(16)
                                        XRASG = "NO"

                                        XRLVL = 9
                                        XRSCD = "0"

                                        If XRTIPO = "USR" Or XRTIPO = "RSP" Then
                                            XRASG = XRCOD
                                            If XRTIPO = "RSP" Then
                                                XRTIPO = "USR"
                                                For YYY = 1 To ATABLA(0, 0)

                                                    If ATABLA(YYY, 0) = XRCOD Then
                                                        XRCOD = ATABLA(YYY, 1)
                                                        XRASG = XRCOD
                                                        XRLVL = ATABLA(YYY, 4)
                                                        XRSCD = ATABLA(YYY, 5)
                                                        Exit For
                                                    End If

                                                Next

                                            End If

                                       Else

									   	XRTIPO = ATABLA(1, 2)
									   	XRCOD = ATABLA(1, 1)
									   	XRASG = ATABLA(1, 1)
									   	XRLVL = ATABLA(1, 4)
									   	XRSCD = ATABLA(1, 5)

									  End If

									        CKTTIPO = XRTIPO
									   		CKTCOD = XRCOD
									   		CKTUSR = XRASG

									   		SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
                                            SINPUT3 = GETRESPOCKT(CKTUSR, "USR")


                                    Else

                                        XRTIPO = ATABLA(2, 2)
                                        XRCOD = ATABLA(2, 1)
                                        XRASG = "NO"
                                        If XRTIPO = "USR" Then
                                            XRASG = XRCOD
                                        End If
                                        XRLVL = ATABLA(2, 4)
                                        XRSCD = ATABLA(2, 5)


                                    End If


                                    CKTTIPO = XRTIPO
                                    CKTCOD = XRCOD
                                    CKTUSR = XRASG





									'if session("APLMDCLI") = "BGAL" and session("TCliente") = "SN" and session("WKF") = "WKF00016631121722600921928PEDSOL" then
									'	CKTCOD = getOwner(session("SEGCKT"))
									'	CKTUSR = CKTCOD
									'end if

                                    SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
                                    SINPUT3 = GETRESPOCKT(CKTUSR, "USR")
                                Else
                                    sinput = "S/D"
                                    SINPUT2 = "S/D"
                                    SINPUT3 = "S/D"
                                End If

                            ElseIf itabs < itab Then
                                If itabs <> 0 Then

                                    sinput = RS2(6) & "." & itabs + 1 & ":" & RS2(7)

                                    XRTIPO = ATABLA(itabs + 3, 2)
                                    XRCOD = ATABLA(itabs + 3, 1)
                                    XRASG = "NO"
                                    XRLVL = ATABLA(itabs + 3, 4)
                                    XRSCD = ATABLA(itabs + 3, 5)


                                    CKTTIPO = XRTIPO
                                    CKTCOD = XRCOD
                                    CKTUSR = XRASG

									'if session("APLMDCLI") = "BGAL" and session("TCliente") = "SN" and session("WKF") = "WKF00016631121722600921928PEDSOL" then
									'	if CKTUSR = "NO" then CKTUSR = CKTCOD
									'end if

                                    SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
                                    SINPUT3 = GETRESPOCKT(CKTUSR, "USR")


                                Else
                                    Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(12) & "' "
                                    Set RS3 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                                    If Not RS3.EOF Then
                                        sinput = RS3(6) & ":" & RS3(7)



                                        If pscclase = "E" Then

                                            XRTIPO = RS3(15)
                                            XRCOD = RS3(16)
                                            XRASG = "NO"

                                            XRLVL = 9
                                            XRSCD = "0"

                                            If XRTIPO = "USR" Or XRTIPO = "RSP" Then
                                                XRASG = XRCOD
                                                If XRTIPO = "RSP" Then
                                                    XRTIPO = "USR"
                                                    For YYY = 1 To ATABLA(0, 0)

                                                        If ATABLA(YYY, 0) = XRCOD Then
                                                            XRCOD = ATABLA(YYY, 1)
                                                            XRASG = XRCOD
                                                            XRLVL = ATABLA(YYY, 4)
                                                            XRSCD = ATABLA(YYY, 5)
                                                            Exit For
                                                        End If

                                                    Next

                                                End If

                                            Else

												XRTIPO = ATABLA(1, 2)
												XRCOD = ATABLA(1, 1)
												XRASG = ATABLA(1, 1)
												XRLVL = ATABLA(1, 4)
	                            				XRSCD = ATABLA(1, 5)

                                            End If

                                             CKTTIPO = XRTIPO
											 CKTCOD = XRCOD
											 CKTUSR = XRASG

											SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
                                            SINPUT3 = GETRESPOCKT(CKTUSR, "USR")

                                        Else

                                            XRTIPO = ATABLA(2, 2)
                                            XRCOD = ATABLA(2, 1)
                                            XRASG = "NO"
                                            If XRTIPO = "USR" Then
                                                XRASG = XRCOD
                                            End If
                                            XRLVL = ATABLA(2, 4)
                                            XRSCD = ATABLA(2, 5)


                                        End If

                                        CKTTIPO = XRTIPO
                                        CKTCOD = XRCOD
                                        CKTUSR = XRASG

                                        SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
                                        SINPUT3 = GETRESPOCKT(CKTUSR, "USR")
                                    Else
                                        sinput = "S/D"
                                        SINPUT2 = "S/D"
                                        SINPUT3 = "S/D"
                                    End If

                                End If
                            End If


                        End If

                        celdas1 = ofv.GenCelda("", "vt style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00cc00;' ", "center", "50%", "", "", "", "<B>SIGUE A PASO:<br> <big>" & sinput & "</big></B>", "si")
                        celdas2 = ofv.GenCelda("", "vt", "", "", "", "", "", "<b>" & sImagene & "</b> " & "<B>DESTINO: " & SINPUT2 & "</B>", "si")
                        celdas3 = ofv.GenCelda("", "vt", "", "", "", "", "", "<b>" & sImageni & "</b> " & "<B>USUARIO: " & SINPUT3 & "</B>", "si")
                        celdas4 = ofv.GenCelda("", "vt", "", "", "", "", "", "&nbsp;", "si")

                        If InStr(session("SGU_nombre"), "NO DISPONIBLE") > 0 Then

                            sinput = "<b>&nbsp;VALIDACION NO DISPONIBLE&nbsp;</b>"

                        Else
            if sacc= "ABR" THEN
                          WSPFRMN = "DTTEMAIT56"
                          sinput = "<a  class=cc href='javascript:void(0);' onClick='this.style.display=""none"";document." & WSPFRMN &".submit();' "
                           IF PSCMOTDAT="R" THEN
                            sinput = sinput & " ><b>&nbsp;REENVIA&nbsp;</b></a>"
                            ELSE
                             sinput = sinput & " ><b>&nbsp;APRUEBA&nbsp;</b></a>"
                           END IF


					     WSPFRMA="cowkfselect.asp?OPCION=OK&" & request.QUERYSTRING
						 sinput = sinput & "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='WK4' style='display:none;' METHOD='POST' ></FORM>"
                         ELSE

                            sinput = "<a  class=cc href=cowkfselect.asp?OPCION=OK&" & request.QUERYSTRING
                            sinput = sinput & " ><b>APRUEBA</b></a>"


                        End If
                        End If

						if btnAP = "N" then
							sInput = ""
	                    	sImagenA = ""
	                    end if

	                    celdas5 = ofv.GenCelda("", "vt align=center", "", "", "", "", "", "<b>" & sImagenA & "</b> " & "<BIG>" & sinput & "</BIG>", "si")

                    Else
                        '''   response.write "paso 2<br>"

                        If RS2(11) Then
                            '''''''''''


                            If pscclase = "E" Then

                                '     response.write "- rs3(15) y (16):" & strsql & "<br>"
                                '     response.write " rs2(11):" & rs2(11) & "- rs3(15) y (16):" & RS3(15) & "-" & RS3(16) & "<br>"

                                XRTIPO = RS2(15)
                                XRCOD = RS2(16)
                                XRASG = "NO"

                                XRLVL = 9
                                XRSCD = "0"

                                If XRTIPO = "USR" Or XRTIPO = "RSP" Then
                                    XRASG = XRCOD
                                    If XRTIPO = "RSP" Then
                                        XRTIPO = "USR"
                                        For YYY = 1 To ATABLA(0, 0)

                                            If ATABLA(YYY, 0) = XRCOD Then
                                                XRCOD = ATABLA(YYY, 1)
                                                XRASG = XRCOD
                                                XRLVL = ATABLA(YYY, 4)
                                                XRSCD = ATABLA(YYY, 5)
                                                Exit For
                                            End If

                                        Next

                                    End If
                                End If

                            Else

                                XRTIPO = ATABLA(2, 2)
                                XRCOD = ATABLA(2, 1)
                                XRASG = "NO"
                                If XRTIPO = "USR" Then
                                    XRASG = XRCOD
                                End If
                                XRLVL = ATABLA(2, 4)
                                XRSCD = ATABLA(2, 5)


                            End If





                            SINPUT2 = GETRESPOCKT(XRCOD, XRTIPO)
                            SINPUT3 = GETRESPOCKT(XRASG, "USR")



                            '''''''''''
                            celdas1 = ofv.GenCelda("", "vt style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00cc00;' ", "center", "50%", "", "", "", "<B>SIGUE A PASO:<br> <big>&nbsp;FIN&nbsp;DEL&nbsp;CIRCUITO</big></B>", "si")
                            celdas2 = ofv.GenCelda("", "vt", "", "", "", "", "", "<b>" & sImagene & "</b> " & "<B>DESTINO: " & SINPUT2 & "</B>", "si")
                            celdas3 = ofv.GenCelda("", "vt", "", "", "", "", "", "<b>" & sImageni & "</b> " & "<B>USUARIO: " & SINPUT3 & "</B>", "si")
                            celdas4 = ofv.GenCelda("", "vt", "", "", "", "", "", "&nbsp;", "si")






                            If InStr(session("SGU_nombre"), "NO DISPONIBLE") > 0 Then

                                sinput = "<b>&nbsp;VALIDACION NO DISPONIBLE&nbsp;</b>"

                            Else
                              '  sinput = "<a  class=cc href=cowkfselect.asp?ACC="&SACC&"&OPCION=OK"
                              '  sinput = sinput & " ><b>&nbsp;APRUEBA&nbsp;</b></a>"
                                if sacc= "ABR" THEN
									WSPFRMN = "DTTEMAIT56"
									sinput = "<a  class=cc href='javascript:void(0);' onClick='this.style.display=""none"";document." & WSPFRMN &".submit();' "

									sinput = sinput & " ><b>&nbsp;APRUEBA&nbsp;</b></a>"



								 WSPFRMA="cowkfselect.asp?ACC=" & SACC & "&OPCION=OK"
								 sinput = sinput & "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='WK4' style='display:none;' METHOD='POST' ></FORM>"
								   ELSE

                                sinput = "<a  class=cc href=cowkfselect.asp?OPCION=OK"
                                sinput = sinput & " ><b>&nbsp;APRUEBA&nbsp;</b></a>"

                            End If
                            End If

                            celdas5 = ofv.GenCelda("", "vt", "center", "", "50", "", "", "<b>" & sImagenA & "</b> " & "<BIG>" & sinput & "</BIG>", "si")

                        Else

                            celdas1 = ofv.GenCelda("", "", "", "50%", "", "", "", "&nbsp;", "si")
                            celdas2 = ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                            celdas3 = ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                            celdas4 = ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                            celdas5 = ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")

                        End If

                    End If

                    If RS2(10) And RS2(13) <> "" Then

                        Strsql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(13) & "' "
                        Set RS3 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
                        If Not RS3.EOF Then
                            sinput = RS3(6) & ":" & RS3(7)





                            If pscclase = "E" Then

                                XRTIPO = RS3(15)
                                XRCOD = RS3(16)
                                XRASG = "NO"

                                XRLVL = 9
                                XRSCD = "0"

                                If XRTIPO = "USR" Or XRTIPO = "RSP" Then

                                    XRASG = XRCOD

                                    If XRTIPO = "RSP" Then

                                        XRTIPO = "USR"
                                        For YYY = 1 To ATABLA(0, 0)

                                            If ATABLA(YYY, 0) = XRCOD Then

                                                XRCOD = ATABLA(YYY, 1)
                                                XRASG = XRCOD
                                                XRLVL = ATABLA(YYY, 4)
                                                XRSCD = ATABLA(YYY, 5)
                                                Exit For
                                            End If

                                        Next

                                    End If

                                Else

									XRTIPO = ATABLA(1, 2)
									XRCOD = ATABLA(1, 1)
									XRASG = ATABLA(1, 1)
									XRLVL = ATABLA(1, 4)
                                	XRSCD = ATABLA(1, 5)

                                End If

                            Else

                                XRTIPO = ATABLA(1, 2)
                                XRCOD = ATABLA(1, 1)
                                XRASG = ATABLA(1, 1)
                                XRLVL = ATABLA(1, 4)
                                XRSCD = ATABLA(1, 5)


                            End If

                            CKTTIPO = XRTIPO
                            CKTCOD = XRCOD
                            CKTUSR = XRASG

                            SINPUT2 = GETRESPOCKT(CKTCOD, CKTTIPO)
                            SINPUT3 = GETRESPOCKT(CKTUSR, "USR")

                        Else
                            sinput = "S/D"
                            SINPUT2 = "S/D"
                            SINPUT3 = "S/D"
                        End If

						'controla q no salga detalles del caso rechaza.
						'puede = 1
                        'if session("APLMDCLI") = "BGAL" and session("TCliente") = "SN" and session("WKF") = "WKF00016631121722600921928PEDSOL" then

						'    sql = "select codigo from segcktresp where segckt='" & session("segckt") & "' and orden <> '999' order by id desc"
						'	set rz = ofv.crearconsultaEx(sql,cn,1,parametros)

						'	if not rz.eof then
						'		if session("USUARIO") = rz(0) then
						'			puede = 0
						'		end if
						'	end if

						'end if

						'if puede = 1 then

                        	celdas1 = celdas1 & ofv.GenCelda("", "vT colspan=2 style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#cc0000;' ", "center", "", "", "", "", "<B>SIGUE A PASO:<br> <big>" & sinput & "</big></B>", "si")
                        	celdas2 = celdas2 & ofv.GenCelda("", "vt colspan=2", "", "", "", "", "", "<b>" & sImagene & "</b> " & "<B>DESTINO: " & SINPUT2 & "</B>", "si")
                        	celdas3 = celdas3 & ofv.GenCelda("", "vt colspan=2", "", "", "", "", "", "<b>" & sImageni & "</b> " & "<B>USUARIO: " & SINPUT3 & "</B>", "si")
                        	celdas4 = celdas4 & ofv.GenCelda("", "vt colspan=2", "", "", "", "", "", "&nbsp;", "si")

						'end if

                        If InStr(session("SGU_nombre"), "NO DISPONIBLE") > 0 Then

                            sinput = "<b>&nbsp;VALIDACION NO DISPONIBLE&nbsp;</b>"

                        Else
                              if sacc= "ABR" THEN
								  WSPFRMN = "DTTEMAIT56R"
								  sinput = "<a  class=cc href='javascript:void(0);' onClick='this.style.display=""none"";document." & WSPFRMN &".submit();' "
								  IF PSCMOTDAT="R" THEN
								   sinput = sinput & " ><b>&nbsp;CANCELA&nbsp;</b></a>"
								  ELSE
									 'puede = 1
									 sinput = sinput & " ><b>&nbsp;RECHAZA&nbsp;</b></a>"
								  	 'if session("APLMDCLI") = "BGAL" and session("TCliente") = "SN" and session("WKF") = "WKF00016631121722600921928PEDSOL" then

									 '   sql = "select codigo from segcktresp where segckt='" & session("segckt") & "' and orden <> '999' order by id desc"
									'	set rz = ofv.crearconsultaEx(sql,cn,1,parametros)

									'	if not rz.eof then
									'		if session("USUARIO") = rz(0) then
									'			puede = 0
									'		end if
									'	end if

									 'end if

								  END IF



								 WSPFRMA="cowkfselect.asp?OPCION=RCH&" & request.QUERYSTRING
								 sinput = sinput & "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='WK4' style='display:none;' METHOD='POST' ></FORM>"
								 ELSE
							'puede = 1
                            sinput = "<a  class=cc href=cowkfselect.asp?OPCION=RCH&" & request.QUERYSTRING
                            sinput = sinput & " ><b>RECHAZA</b></a>"

                        End If
                        End If

						'if puede = 1 then
                        	celdas5 = celdas5 & ofv.GenCelda("", "vt align=center", "", "", "", "", "", "<b>" & sImagenR & "</b> " & "<BIG>" & sinput & "</BIG>", "si")
						'else
							celdas5 = celdas5 & ofv.GenCelda("", "vt align=center", "", "", "", "", "", "&nbsp;", "si")
						'end if

                    Else

                        celdas1 = celdas1 & ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                        celdas2 = celdas2 & ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                        celdas3 = celdas3 & ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                        celdas4 = celdas4 & ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                        celdas5 = celdas5 & ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")

                    End If
                     ' response.write session("SegGrpLst") & "222222" & Session("SegGrp")
                      if session("USRPRV")  = "I" then
                      if trim(session("SegGrpLst"))="" then session("SegGrpLst")=Session("SegGrp")
                      end if


                       XMFTGRP = split(session("SegGrpLst"),chr(9))
                       XFStrSql=""
                        for XMFI = 0 to ubound(XMFTGRP)
					         XFStrSql  = XFStrSql  & XMFSEP & "'" & replace(XMFTGRP(XMFI) & "","'","''") & "'"
					         XMFSEP  = ","
                        next
'responm

				'	sql = "select grupo from usuarios a inner join GruposAprob b on (a.Grupo=b.mapgrp) where a.id='" & session("Usuario") & "'"

					sql = "select mapgrp from GruposAprob where estado=1 and  mapgrp in(" & XFStrSql & ")"
					'response.write sql
					Set rt = ofv.crearconsultaex(sql,cn2,1,parametros)

					strsql = "select ordenest from objwfl where codobjeto=" & nCodigo & " and tipobjeto='" & stipo & "'"
				'	response.write strsql
					set rx = ofv.crearconsultaex(strsql,cn2,1,parametros)
					sorden=""
					if not rx.eof then sorden = rx(0)

					If Not rt.eof Then

						if session("WKF") = "WKF00016631121722600921928DD" then

							sImage = "<font face='corflaticon' size=3 color='#FE9A2E;'>q</font>"
							shref = "wkfreasignador.asp?ID=" & nCodigo & "&tipobjeto=" & stipo & "&WKFCOD=" & session("SEGCKTPSC") & "&OPWK=REASIGNAR" & "&CantRegMoverx=" & CantRegMoverx
							sinput = "<a  class=cc href=" & shref & ">"
							sinput = sinput & sImage & "<b>REASIGNAR</b></a>"
							celdas6 = ofv.GenCelda("", "vt colspan=2 align=center", "", "", "", "", "","<BIG>" & sinput & "</BIG>", "si")

						end if

					Else

						celdas6 = ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")

					End If

					Call ofv.cerrarconn(rt)



                    filas = ofv.GenRow("", "", "", "", "10", "", "", celdas1, "si")
                    filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas2, "si")
                    filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas3, "si")
                    filas = filas & ofv.GenRow("", "", "", "", "10", "", "", celdas4, "si")
                    filas = filas & ofv.GenRow("", "", "", "", "", "", "", celdas5, "si")
                    filas = filas & ofv.GenRow("", "", "", "", "", "", "", celdas6, "si")

                    sHTML = sHTML & ofv.GenTabla("", "aaa ", "", "100%", "", "0", "", "-2", "-1", "", "", filas, "si")

                End If

            End If

        End If

    End If 'finaliza el OK o Rechazado




IF SACC="ABR" THEN

    sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%   ","","0","","1","0","","",Filann,"si")
    sHTML = sHTML & "<form name=initq action='NOVEDMENU.asp' "
    sHTML = sHTML & " method=post style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

ELSE

	celdax = ofv.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
	filax = ofv.GenRow("", "", "", "", "", "", "", celdax, "si")
	filax = filax & ofv.GenRow("", "", "", "", "", "", "", celdax, "si")
  shref = "dictamenes.asp?CODIGO=" & session("NCliente") & "&TIPO=" & session("TCliente")
  sInputx = "<a style='color:#000000;' href= " & shref & " onclick=""window.open(this.href,this.target,'width=600,height=300,top=200,left=350,resizable=yes,scrollbars=yes');return false;"">DICTÁMENES</a>"
  celdax = ofv.GenCelda("", "ttr style='background-color:#FE9A2E;'", "center", "", "", "", "",sInputx, "si")
  filax = filax & ofv.GenRow("", "", "center", "", "", "", "", celdax, "si")
  sHTML = sHTML & ofv.GenTabla("", "aaa", "center", "10%", "", "0", "", "-2", "-1", "", "", filax, "si")

  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"


  Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
  Filann = ofv.GenRow("","","","","","","",Celda,"si")


  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%   ","","0","","1","0","","",Filann,"si")






  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

    'Botonera
    Agregar = ""
    Buscar = ""
    celda = ""

    sHTML = sHTML & ofv.botoneraESSX("0", rs, cantregmoverx, cantregmostrarx, "no", "no", posi, "", "", "no", Agregar, Buscar, celda)
 END IF

    Call ofv.cerrarconsulta(rs)
    Call ofv.cerrarconn(cn)
    if ACT="S" and (sopcion="OK" OR SOPCION="RCH") and BrowserName() <> "XFIREFOX" THEN
     xsfrm="<form id=xx name=xx style='display:none' action='novedmenu.asp?tab=view4' method=post target=info>"
    END IF

    sHTML = sHTML  & xsfrm   & "</body></html>"
End If

	sHTML = sHTML & "<script type='text/javascript'>"
	if ACT="S" and (sopcion="OK" OR SOPCION="RCH") and BrowserName()<> "XFIREFOX" THEN
	 sHTML = sHTML & " document.getElementById('xx').submit();"
	END IF
	sHTML = sHTML & "function disp(){"
	sHTML = sHTML & "  }"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    fr.action=fr.action;"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    fr.action=fr.action;"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"

response.write sHTML
'response.write session("SEGCKTPSC") & "   " & request.QUERYSTRING

'' response.write sopcion & " " & sopx & "<br>"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FUNCIONES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GENCKTPSC(XPSC, XSIG, XMODEL, XFIN, XRTIPO, XRCOD, XRASG, XPDSC, XOBSERV, XNIVEL, XSUBCOD, IDRESP)

'response.write "XPSC:" & XPSC & "<br>" & "XSIG:" & XSIG & "<br>" & "XMODEL:" & XMODEL & "<br>" & "XFIN:" & XFIN & "<br>" & "XRTIPO:" & XRTIPO & "<br>" & "XRCOD:" & XRCOD & "<br>" & "XRASG:" & XRASG & "<br>" & "XPDSC:" & XPDSC & "<br>" & "XOBSERV:" & XOBSERV & "<br>" & "XNIVEL:" & XNIVEL & "<br>" & "XSUBCOD:" & XSUBCOD & "<br>" & "SOPCION: " & SOPCION
    SEGCKTPSC = XPSC
'response.write session("APLMDCLI") & "--" & session("TCliente") & "--" & session("WKF")
'-------------------------------------------------------------------------------
'ESTA RUTINA CIERRA EL PASO 999 AUTOMATICO
    sqlF="select cast ( fin as int)  from WKFPASOS where wkfpascod='" & XMODEL   & "'"
    set rspsfinal=ofv.crearconsultaEx(sqlf, cn, 1, parametros)
    spsfin="N"
    if not rspsfinal.eof then
     if rspsfinal(0) & ""  = "1"  AND session("APLMDCLI") <> "BPN" then spsfin="S"
    end if
    call ofv.cerrarconsulta(rspsfinal)

'----------------------------------------------------------------------------------------

    XCOM = GETDH()
    If XFIN = "SI"  or spsfin="S" Then
        XFIN = GETDH()
    Else
        XFIN = "NO"
    End If




	'inserta paso siguiente
    Strsql = "insert into SEGPASO (SEGCKT,SEGCKTPSC,MODEL,COMIENZO,FIN,OKRCH,RESPTIPO,RESPONSABLE,RESPASG,PSCDSC,OBSERVACIONES,NIVEL,SUBCODIGO,SEGIDRESP) values "
    Strsql = Strsql & " ('" & session("SEGCKT") & "','" & SEGCKTPSC & "','" & XMODEL & "','" & GETDH() & "','" & XFIN & "','" & XSIG & "','" & XRTIPO & "','" & XRCOD & "','" & XRASG & "','" & XPDSC & "','" & XOBSERV & "','" & XNIVEL & "','" & XSUBCOD & "','" & IDRESP & "')"
    Call ofv.crearconsultaEx(Strsql, cn, 1, parametros)

    if spsfin="S" AND session("APLMDCLI") <> "BPN" then
   	  Strsql = "UPDATE SEGCIRCUITO SET FIN = '" & GETDH() & "'  WHERE SEGCKT = '" & session("SEGCKT") & "' "
      Call ofv.crearconsultaEx(Strsql, cn, 1, parametros)
    end if


'controla que tenga algun destino y que no sea 0 el formulario para entrar a la rutina de copiado de documentos sin que pinche
	if session("APLMDCLI") = "BGAL" and session("TCliente") <> "PDCORXY" and session("TCliente") <> "SN" then
     sql = "select iddestino from envioformdet2 where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "' order by iddestino desc"
	 set rk  = ofv.crearconsultaEx(sql,cn2,1,parametros)

	 flg = 0
	if not rk.eof then

		do while not rk.eof

			if rk(0) <> 0 then
				flg = 1
			else
				flg = 0
			end if

		rk.movenext
		loop

	 end if

    'RUTINA PARA COPIAR LOS ADJUNTOS DEL DOCUMENTO A LOS DESTINOS PREVIAMENTE SELECCIONADOS
    If session("COPIA") = "S" and tipocpt = "K" and spsfin = "S" and flg = 1 Then

		xsql = "select estado from objprty where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "'"
		set ri = ofv.crearconsultaEx(xsql,cn2,1,parametros)

		xestado = ""
		if not ri.eof then xestado = ri(0)

    		if xestado = 3 then

				sql = "select c.directorio,c.archivo,d.url from objetos a "
				sql = sql & " left join envioformdet2 b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto) "
				sql = sql & " left join objhtm c on (a.codobjeto=c.codobjeto and a.tipobjeto=c.tipobjeto) "
				sql = sql & " left join destinosurl d on (b.iddestino=d.id) "
				sql = sql & " where a.codobjeto=" & session("NCliente") & " and a.tipobjeto='" & session("TCliente") & "'"
				set ro = ofv.crearconsultaEx(sql,cn2,1,parametros)

				if not ro.eof then

					dim fs
					set fs=Server.CreateObject("Scripting.FileSystemObject")

					do while not ro.eof

						xOrigen = session("DirUpload") & ro(0) & "\" & ro(1)
						xDestino = ro(2)

						fs.CopyFile xOrigen,xDestino

					ro.movenext
					loop

					set fs=nothing

					asunto = "Claridad Galicia - Formulario publicado"

					sql = "select a.url,c.descri,a.mail1,a.mail2 from destinosurl a "
					sql = sql & " inner join envioformdet2 b on (a.id=b.iddestino) "
					sql = sql & " inner join objetos c on (b.codobjeto=c.codobjeto and b.tipobjeto=c.tipobjeto) "
					sql = sql & " where b.codobjeto=" & session("NCliente") & " and b.tipobjeto='" & session("TCliente") & "'"
					set rj = ofv.crearconsultaEx(sql,cn2,1,parametros)

					cuerpo = "A continuación se informan las urls a las que se copio el archivo adjunto del formulario: " & rj(1) & "<br><br>"
					mails = ""
					sep = ""
					do while not rj.eof

						cuerpo = cuerpo & rj(0) & "<br>"
						mails = mails & sep & rj(2) & ";" & rj(3)

						sep = ";"

					rj.movenext
					loop

					'xx = SendMailAtt(cuerpo,mails,"mabraham@corconsulting.com.ar",asunto,"CDOSYS","SI","1","NO","","xClaridadGalicia@bancogalicia.com.ar",0)

					'Se modifica la rutina de envío de mail para que tome por defecto
					'la dirección del sistema en el remitente ("De" o "From") y en el campo "CC" (Con copia)
					xx = SendMailAtt(cuerpo,mails,session("ORGMAIL"),asunto,"CDOSYS","SI","1","NO","",session("ORGMAIL"),0)

				end if

    		end if

    end if

    end if


End Function

Function getOwner(codWkf)

ssql = "select codigo from segcktresp where segckt='" & codWkf &  "' and modo='D'"
Set rsm = ofv.crearconsultaEx(ssql, cn, 1, parametros)

xCod = ""
if not rsm.eof then xCod = rsm(0)

getOwner = xCod

End function


Function GENACCION(XPASO, XTPP)

     'response.write "entro accion: " & XPASO & " -- " & XTPP & " // WKFAUTO" & WKFAUTO
      Set xcn2 = ofv.conectar(ofv.strconn0)

     SQL = "SELECT TIPOCPT,WKFAUTO FROM TIPOBJETO WHERE  TIPOBJETO='" & session("TCliente") & "'"
    ' response.write "SQL:" & sql & "<br>"
     Set RS1 = ofv.crearconsultaEx(SQL, xcn2, 1, parametros)
     PEDIDO="N"
     modelo = rs1(1)
     IF NOT RS1.EOF THEN  PEDIDO=RS1(0)

   ' response.write "pedido " & PEDIDO & "<BR>"
     CALL OFV.CERRARCONSULTA(RS1)

     call ofv.cerrarconn(xcn2)


    Strsql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & XPASO & "' AND (TIPO = '" & XTPP & "' OR TIPO = 'FIN') "
    Set RS1 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)

    estado = RS1(5)

    If Not RS1.EOF Then
        If RS1(4) = "A" Then
         '      response.write "estado: " & RS1(5) & "<br>"
           if rs1(0) & "" ="WKF00016631121722600921928PED" OR PEDIDO="PD" OR PEDIDO="SN" then
             GETWKFTAB="PEDIDOS"
           else
            If session("TCliente") = "RSK" Then
                GETWKFTAB = "ORRSKCIRCUITO"
            Else
                GETWKFTAB = "OBJPRTY"
            End If
           end IF

			IF WKFAUTO = "WKF00016631121722600921928DD" and SOPCION = "RCH" then estado = 1

            xxxx = GETWKFMANAGER(session("NCliente"), session("TCliente"), estado, GETWKFTAB, "WKF")

        End If


    End If




End Function





Function GENCKTMAIL(WETIPO, WECOD, WEMESSAGE, WEORIGEN, WELVL)


    XRESPO = WECOD & ""

    If XRESPO <> "" And XRESPO <> "NO" Then

        Set cn2 = ofv.conectar(ofv.strconn0)

        XPARA = GETRESPOCKT(WECOD, WETIPO)
       ' Strsql = "SELECT EMAIL,DESCRIPCION FROM USUARIOS WHERE "

        XGUMENSAGEw = WEMESSAGE

        XMESAURL = ""

        IF Session("CORUNQWKF") = "S" THEN

           XMESAURL = getobjurlWK(session("NCliente"),session("TCliente"),SEGCKTPSC)

        ELSE

           If session("URLTYP") = "V" Then

              XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))

           End If


           If session("URLPORT") = "N" Then
              XMESAURL = TipoSite("") & request.servervariables("SERVER_NAME") & XMESAURL
           Else
              XMESAURL = TipoSite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
           End If

          ' XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL


        END IF




        XGUMENSAGEw = Replace(XGUMENSAGEw, "<APLURL>", "<U><A HREF='" & XMESAURL & "' TARGET='_blank' >")
        XGUMENSAGEw = Replace(XGUMENSAGEw, "</APLURL>", "</A></U>")

        WEMESSAGE = XGUMENSAGEw


        Select Case WETIPO

            Case "USR"
                'STRSQL = STRSQL & " ID = '" & XRESPO & "' "
            SELECT CASE  session("APLMDCLI")
                 CASE "GP"

				  IdDomE=SESSION("CLIEMP")

				   sqlw="select codempre from SEGCKTRESP where id=0" & XRID
				   set rsXTt=ofv.crearconsultaEx(sqlw,cn,1,parametros)
					 if not rsXTt.eof then
						  IF rsXTt("codempre") & "" <>  "" THEN IdDomE= rsXTt("codempre")
					 end if
				   call ofv.cerrarconsulta(rsXTt)

				 xxxx = getInfoUserWeX(XRESPO,IdDomE)
				CASE ELSE
				 xxxx = getInfoUserW(XRESPO)

             END SELECT

                'xxxx = getInfoUserw(XRESPO)


                If session("SGU_mail") <> "" Then sInputx = session("SGU_mail")


                stowho = sInputx & ""
                If Len(stowho) > 0 Then
                    stowho = Trim(stowho)
                    if wkfT="AUTO" THEN

                    subject = "SOLICITUD DE AUTORIZACION"
					ELSE
					  subject = "SOLICITUD DE OBSERVACION"
					END IF
					IF tipocpt="PD" THEN
						subject = "SOLICITUD ABM PEDIDO"
					elseif tipocpt="SN" then
						subject = "SOLICITUD ABM NORMATIVA"
					END IF

					sql = "select responsable from segcircuito where segckt='" & session("segckt") & "'"
					set rp = ofv.crearconsultaex(sql,cn,1,parametros)

					if SOPCION = "RCH" then

						if rp(0) = XRESPO then
							if session("WKF") = "WKF00016631121722600921928SOL" then
								subject = "RECHAZO DE SOLICITUD ABM NORMATIVA"
							else
								subject = "RECHAZO DE AUTORIZACION"
							end if

						else
							subject = "SOLICITUD DE AUTORIZACION"
						end if

					end if


                    If WETIPO = "USR" Then
                        XPARA2 = XPARA
                    Else
                        XPARA2 = XPARA & " / " & session("SGU_nombre")
                    End If
                    ''  message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
                    Xmessage = GenMDLMAIL(XPARA2, WEORIGEN, WEMESSAG )
                    Call SendMail(Xmessage, stowho, WEORIGEN, subject, session("MailComponent"), "SI", "1", "NO")
                End If


            Case "GRP"

                'STRSQL = STRSQL & " GRUPO = '" & XRESPO & "' and ( NOT nivel >  " & WELVL & " ) "

                Call getUsrIngroup(XRESPO)
                A_SGU_UIGuserid = Split(session("SGU_UIGuserid"), Chr(9))
                A_SGU_UIGdescri = Split(session("SGU_UIGdescri"), Chr(9))

                '   set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                ' DO UNTIL RSXPA.EOF
                For a_i = 0 To UBound(SGU_UIGdescri) - 1
                    xxxx = getInfoUserw(A_SGU_UIGuserid(a_i))
                    stowho = session("SGU_mail") & ""
                    If Len(stowho) > 0 Then
                        stowho = Trim(stowho)
                        if wkft="AUTO" THEN

                        subject = "SOLICITUD DE AUTORIZACION"
					   	ELSE
					   		subject = "SOLICITUD DE OBSERVACION"
					    END IF

					    IF tipocpt="PD" THEN
						  subject = "SOLICITUD ABM PEDIDO"
						END IF
                        If WETIPO = "USR" Then
                            XPARA2 = XPARA
                        Else
                            XPARA2 = XPARA & " / " & A_SGU_UIGdescri(a_i)
                        End If
                        ''' message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
                        Xmessage = GenMDLMAIL(XPARA2, WEORIGEN, WEMESSAGE)
                        Call SendMail(Xmessage, stowho, WEORIGEN, subject, session("MailComponent"), "SI", "1", "NO")
                    End If

                    'RSXPA.MOVENEXT
                Next



            Case "SEC"
                XRESPO = 0 & XRESPO

                Strsql = Strsql & " SECTOR = " & XRESPO & " and ( NOT nivel >  " & WELVL & " ) "

        End Select






    End If

End Function


Function GETRESPOCKT(RESPCOD, RESPTIP)

    GETRESPOCKT = "SIN ASIGNAR"
    sInputxEco = "SIN ASIGNAR"



    XRESPO = RESPCOD & ""

    If XRESPO <> "" And XRESPO <> "NO" Then

        Set cn2 = ofv.conectar(ofv.strconn0)
        Set cn3 = ofv.conectar(ofv.strconn2)

        Select Case RESPTIP

            Case "USR"

                xxxx = getInfoUserw(XRESPO)
                sInputx = session("SGU_nombre")

                ' strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
                ' set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                ' if not rsniv.eof then sInputx = rsniv(1)
                ' call ofv.cerrarconsulta(rsniv)
                GETRESPOCKT = sInputx

            Case "GRP"
                Strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
                Set rsniv = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
                If Not rsniv.EOF Then sInputxEco = rsniv(1)
                Call ofv.cerrarconsulta(rsniv)
                GETRESPOCKT = sInputxEco

            Case "SEC"
                XRESPO = 0 & XRESPO
                Strsql = "Select * From ECO where ECO = " & XRESPO
                '''     response.write strsql
                Set rsniv = ofv.crearconsultaEx(Strsql, cn3, 1, parametros)
                If Not rsniv.EOF Then sInputxEco = rsniv(1)
                Call ofv.cerrarconsulta(rsniv)
                GETRESPOCKT = sInputxEco
        End Select

        '  GETRESPOCKT = sInputx

    End If


End Function

Function GETDH()

    XKD = DatePart("d", Date)
    XKM = DatePart("m", Date)
    XKY = DatePart("yyyy", Date)





    GETDH = XKD & "/" & XKM & "/" & XKY & " " & Time()


End Function


Function GetProceso(xcodigo)


    xdescri = ""
    sinput = ""
    Strsql = ""
    Strsql = Strsql & " select x.periodo, A.*, W.ESTADO "
    Strsql = Strsql & " from csrelpathdescri a   "
    Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "
    Strsql = Strsql & " inner join orrskcircuito w on (a.cobj = w.codobjeto)   "
    Strsql = Strsql & " inner join orperiodos x on (w.periodo = x.objid )  "
    Strsql = Strsql & " where  a.cobj = 0" & xcodigo & "   "

    '
    '     response.write strsql & chr(13)


    Set rsX1 = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
    If Not rsX1.EOF Then

        For WKI = 1 To 8

            WKIPROC = rsX1(WKI) & ""

            If WKIPROC <> "" Then
                If WKIPROCD = "" Then WKIPROC = WKIPROC
                WKIPROCD = WKIPROCD & WKIPROC & " / "
            End If


        Next

        If Len(WKIPROCD) > 4 Then
            WKIPROCD = Left(WKIPROCD, Len(WKIPROCD) - 3)
        End If


        sinput = WKIPROCD & " - (" & rsX1(0) & ")"

    End If
    Call ofv.cerrarconsulta(rsX1)
    GetProceso = sinput
End Function

Function tieneDesa(cod,tip)

vSql = "select objid from objwfldet where codobjeto=" & cod & " and tipobjeto='" & tip & "'"
set td = ofv.crearconsultaEx(vSql,cn2,1,parametros)

sTiene = 0
if not td.eof then sTiene = 1

tieneDesa = sTiene

End Function

%>
