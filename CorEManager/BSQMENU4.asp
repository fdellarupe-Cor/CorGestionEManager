<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/novedformat.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"
   Dim ANTabla(500, 25)
   moverreg = clng("0" & request.querystring("canTregmover"))
   tope=clng("0" & request.querystring("tope"))
   MAINAPL=request.querystring("MAINAPL")

  nFecini = 0

      If nFecini < CLng(CDate("01-01-1900")) Then
          nFecini = CLng(CDate("01-01-1900"))
      Else
          nFecini = 0
      End If

    dFechai = CLng(Date) - nFecini

     WSPDIVN = "DTNOVINT4"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
     sHTML = ""
     sHTML = sHTML & "<HTML>"
     sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"


''''''''''''''




    Set cn = OFV.conectar(OFV.strconn2)

                 Fila    = ""


                SESSION("Segce") = "S"
                SESSION("Segco") = "S"
                cuenta=0

                StrSql = "SELECT * FROM csnovedades "
                StrSql = StrSql & " Where fecexp > " & dFechai
                StrSql = StrSql & " and fecinicio < " & dFechai + 1
                StrSql = StrSql & " And tipo = 4"
                StrSql = StrSql & " And usugrp = '" & SESSION("usuario") & "' order by fecinicio desc, importancia, objid desc"
                'RESPONSE.WRITE STRSQL
                Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If rsx.EOF Then
                    xnohay1 = "S"
                Else
                    xnohay1 = "N"
                    if Not rsx.EOF then rs.move(moverreg)
                    do WHILE (NOT  rsx.eof) AND (cuenta < tope)
                        cuenta=cuenta+1

                        INTabla = INTabla + 1

                        StrSql = "SELECT archivo FROM tipobjeto "
                        StrSql = StrSql & " Where tipobjeto = '" & rsx(9) & "'"
                        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                        If Not rs.EOF Then
                            xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
                        Else
                            xarc = "NO"
                        End If
                        Call ofv.cerrarconsulta(rs)


                        SESSION("ASP") = "novedmenu.asp?menu=4"
                        SESSION("sEstado") = "PARA APROBAR"
                        SESSION("sEstadon") = "2"
                        SESSION("Form2") = "apr"

                        aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&Perfil=xxxx&Funcion=E"
                        aspeditar = CStr(aspeditar)
                        sHref = "&nbsp;"
                        ' response.write sActcon
                        sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                   '''''wNomL = "&nbsp;" & Capitalize(rsx(5)) & "&nbsp;"
                        wNomL = "&nbsp;" & rsx(5) & "&nbsp;"

                        sImagen = wNomL


                        If xarc = "NO" Then
                            xarc = "&nbsp;"
                        End If

                        ''''  sLink = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=" & xscolor & " >" & sImagen & "</font>", "si")


                        sLink = sImagen
                        'XTSLNKM = OFV.GenLink("", "cmb ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                        XTSLNKM = ofv.GenLink("", "istt ", sHref, "", "", sFuncion, "XXMSGXX", "si")


                        xfromusr = rsx(11)


                        xxxx = getInfoUser(rsx(11))
                        xfromusr = SESSION("SGU_nombre")

                        ANTabla(INTabla, 0) = rsx(1)
                        ANTabla(INTabla, 1) = 4
                        ANTabla(INTabla, 2) = CLng(rsx(6))
                        ANTabla(INTabla, 3) = CLng(rsx(3))
                        ANTabla(INTabla, 4) = rsx(12)
                        ANTabla(INTabla, 5) = "S"
                        ANTabla(INTabla, 6) = "N"
                        ANTabla(INTabla, 7) = xfromusr
                        ANTabla(INTabla, 8) = sLink
                        ANTabla(INTabla, 9) = xarc
                        ANTabla(INTabla, 10) = XTSLNKM
                        ANTabla(INTabla, 11) = xmenu
                        ANTabla(INTabla, 12) = rsx(13)
                        ANTabla(INTabla, 13) = rsx(14)
                        ANTabla(INTabla, 14) = rsx(15)
                        ANTabla(INTabla, 15) = rsx(16)
                        ANTabla(INTabla, 16) = rsx(18)

                        ANTabla(INTabla, 17) = rsx(19) & ""
                        ANTabla(INTabla, 18) = rsx(20) & ""
                        ANTabla(INTabla, 19) = rsx(21) & ""
                        ANTabla(INTabla, 20) = rsx(22) & ""

                        ANTabla(INTabla, 21) = "0"




                        rsx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)

                '''''''''''''


                Set cn1 = ofv.conectar(ofv.strconn5)
                Set cn4 = ofv.conectar(ofv.strconn6)


                If wkaplic = "OR" Or wkaplic = "PL" Then

                    xtipos = " OBJTIPO = 'RSK' "


                Else

                    xtipos = " OBJTIPO in ("
                    separ = ""

                    StrSql = "SELECT tipobjeto FROM tipobjeto "
                    Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs.EOF Then
                        Do While Not rs.EOF
                            xtipos = xtipos & separ & "'" & rs(0) & "'"
                            separ = ","
                            rs.movenext
                        Loop
                    Else
                        xtipos = xtipos & "' '"
                    End If
                    xtipos = xtipos & ")"

                End If





                SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & SESSION("Usuario") & "') "


                xxxx = getInfoUser(SESSION("Usuario"))
                xusrlog = SESSION("SGU_nombre")
                'RESPONSE.WRITE xusrlog

                If xusrlog = "NO" And SESSION("USRPRV") = "LACAJA" Then
                    SESSION("USRPRV") = "I"
                    SESSION("Usuario") = "USRCONSULTA"
                    Call getInfoUser(SESSION("Usuario"))
                   xusrlog = SESSION("SGU_nombre")
                   SESSION("USRPRV") = "LACAJA"
                End If

               ' xxxx = getInfoUser(session("Usuario"))
               ' xusrlog = session("SGU_nombre")


              '  RESPONSE.WRITE xusrlog &" 222222222222222222222222222" & CHR(13)
                If xusrlog <> "NO" Then


                    SEGWHERE = SEGWHERE & "OR ( ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & SESSION("Usuario") & "') ) "
                    SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

                    '''''''''''''''''''''''''''''''''''''''''''
                    If SESSION("DELLOG") = "S" Then

                        XMFTGRP = Split(SESSION("SegGrpLst"), Chr(9))
                        xcount = UBound(XMFTGRP) + 1
                        XMFSEP = ""
                        SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

                        For XMFI = 0 To UBound(XMFTGRP)
                            SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                            XMFSEP = ","
                        Next

                        SEGWHERE = SEGWHERE & ") ) )"

                    Else
                        SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
                    End If


                    ''''''''''''''''''''''''''''''''''''''''''

                    '''   IF  RS(21) > 0 THEN


                    If CStr(SESSION("SGU_idcostcenter") & "") <> "0" Then
                        SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                        SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & SESSION("SGU_idcostcenter") & "' OR SUBCODIGO = '" & SESSION("SGU_idcostcenter") & "' ) "
                        SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
                        SEGWHERE = SEGWHERE & " ) ) "
                    End If


                End If
                If xusrlog <> "NO" Then
                  SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "
                Else
                  SEGWHERE = SEGWHERE & " ) ORDER BY PSCCOMIENZO "
                End If

                StrSql = "SELECT * FROM CKTVIEW " & SEGWHERE

                'RESPONSE.WRITE STRSQL & "<BR>" & chr(13)
                Set rsx = ofv.crearconsultaEx(StrSql, cn1, 1, parametros)
                If rsx.EOF Then
                    xnohay2 = "S"
                Else

                    Do Until rsx.EOF

                        xvale = "N"

                        If rsx(3) = "RSK" Then

                            StrSql = "SELECT * FROM ORRSKCIRCUITO "
                            StrSql = StrSql & " Where OBJID = 0" & rsx(4)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)


                        Else



                            StrSql = "SELECT * FROM objetos "
                            StrSql = StrSql & " Where codobjeto = 0" & rsx(4)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)

                        End If




                        If xvale = "S" Then

                            xnohay2 = "N"
                            cuenta=cuenta +1
                            INTabla = INTabla + 1


                            xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

                            SESSION("ASP") = "novedmenu.asp?menu=4"
                            SESSION("sEstado") = "PARA APROBAR"
                            SESSION("sEstadon") = "2"
                            SESSION("Form2") = "apr"

                            aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL


                            ''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

                            aspeditar = CStr(aspeditar)
                            sHref = "&nbsp;"
                            '        response.write sActcon
                            sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                            wNomL = "" & Capitalize(rsx(7)) & "&nbsp;-&nbsp;Paso&nbsp;" & Capitalize(rsx(16)) & "&nbsp;"
                            wNomL = "" & rsx(7) & "&nbsp;-&nbsp;Paso&nbsp;" & rsx(16) & "&nbsp;"

                            sImagen = wNomL


                            sImagen = sImagen & "-&nbsp;" & EsqnomP(rsx(3), rsx(4), ofv)

                            If xarc = "NO" Then
                                xarc = "&nbsp;"
                            End If

                            sLink = sImagen
                            'XTSLNKM = OFV.GenLink("", "cmb title='Autorizar documento ...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                            XTSLNKM = ofv.GenLink("", "istt title='Autorizar documento ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

                            xfromusr = rsx(6)


                            xxxx = getInfoUser(rsx(6))
                            xfromusr = SESSION("SGU_nombre")

                            ANTabla(INTabla, 0) = "CKT" & INTabla
                            ANTabla(INTabla, 1) = 5
                            ANTabla(INTabla, 2) = 0
                            ANTabla(INTabla, 3) = dFechai
                            ANTabla(INTabla, 4) = rsx(10)
                            ANTabla(INTabla, 5) = "S"
                            ANTabla(INTabla, 6) = "N"
                            ANTabla(INTabla, 7) = xfromusr
                            ANTabla(INTabla, 8) = sLink
                            ANTabla(INTabla, 9) = xarc
                            ANTabla(INTabla, 10) = XTSLNKM
                            ANTabla(INTabla, 11) = xmenu


                            ANTabla(INTabla, 12) = 2
                            ANTabla(INTabla, 13) = EsqnomP(rsx(3), rsx(4), ofv)
                            ANTabla(INTabla, 14) = sLink
                            ANTabla(INTabla, 15) = "0"
                            ANTabla(INTabla, 16) = ""

                            ANTabla(INTabla, 17) = ""
                            ANTabla(INTabla, 18) = ""
                            ANTabla(INTabla, 19) = ""
                            ANTabla(INTabla, 20) = ""

                            ANTabla(INTabla, 21) = "0"


                      If cuenta >= tope Then ' CANT REG DE LA NOVEDADES

					   WSPFRMN = "seccion" '& "2222222222"
					   sxlinksig =  "<A id=lnksig class=nt style='display:Xnone;'  href='javascript:void(0)' ONCLICK=document." &WSPFRMN & ".submit(); "
					   sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"

					   sxlinkant =  "<A id=lnkant class=nt style='display:none;' HREF=BSQcMD.asp"

						sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
						botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

					   WSPFRMN = "seccion"
							   'WSPFRMA = "consdocumindex.asp"
					   WSPFRMA = "BSQMENU4.asp?canTregmover=" & cuenta & "&tope=" & tope &  "&MAINAPL=" & xMAINAPL

					   WSPFORMEsecc = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST'></FORM>"


					   INTabla=INTabla+1
					   ANTabla(INTabla, 0) = "0"
					   ANTabla(INTabla, 1) = 13
					   ANTabla(INTabla, 2) = 1
					   ANTabla(INTabla, 3) = dFechai
					   ANTabla(INTabla, 4) = 0
					   ANTabla(INTabla, 5) = "S"
					   ANTabla(INTabla, 6) = "N"
					   ANTabla(INTabla, 7) = "MASREG"
					   ANTabla(INTabla, 8) = ""
					   ANTabla(INTabla, 9) = xarc
					   ANTabla(INTabla, 10) = "<font size=2 color=red face=tahoma>Ver Más</font>" & botones & WSPFORMEsecc
					   ANTabla(INTabla, 11) = xmenu
					   ANTabla(INTabla, 12) = "1"
					   ANTabla(INTabla, 13) = ""
					   ANTabla(INTabla, 14) = ""
					   ANTabla(INTabla, 15) = ""
					   ANTabla(INTabla, 16) = ""
					   ANTabla(INTabla, 17) = ""
					   ANTabla(INTabla, 18) = ""
					   ANTabla(INTabla, 19) = ""
					   ANTabla(INTabla, 20) = ""

					   ANTabla(INTabla, 21) = "0"
					   Exit Do
                        end if

                        End If



                        rsx.movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)

         shtml=shtml & noveformat(4,"111",4)
          if rs.eof then
		    sxlinksig =""
		    else
		    sxlinksig =  "<A id=lnksig class=nt style='display:Xnone;'  HREF='BSQMENU4.asp?CANTREGMOVER=" & moverreg + tope  & "&tope="& tope &"&MAINAPL=" & MAINAPL &"'"
		    sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"
		    end if

		    if moverreg - tope< 0 then
		    sxlinkant=""
		    else
		     sxlinkant =  "<A id=lnkant class=nt style='display:Xnone;' HREF='BSQMENU4.asp?CANTREGMOVER=" &  moverreg - tope  & "&tope="& tope  &"&MAINAPL=" & MAINAPL &"'"

		 	sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
		    end if

		 	botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

		   botones  ="<font size=2 color=red face=tahoma>Ver Más</font>" & botones
    sHTML = sHTML & botones
response.write sHTML & "</body></html>"
end if
%>
