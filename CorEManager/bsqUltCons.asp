<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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

   moverreg = clng("0" & request.querystring("canTregmover"))
   tope=clng("0" & request.querystring("tope"))
   MAINAPL=request.querystring("MAINAPL")



     WSPDIVN = "DTNOVINT5"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
     sHTML = ""
     sHTML = sHTML & "<HTML>"
     sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"


''''''''''''''




    Set cn = OFV.conectar(OFV.strconn2)

  Fila    = ""



   xTrColor = "#FFFFFF"

  stitem = "Tema"
  sIcono = "ICorTema.jpg"
  sNuevoTipo = "SUBTEMAS"
  nNivel = 0

  sID = 0
  nv = 0
  cborderB = Session("CBRDBOT")
  cborderl = Session("CBRDLFT")
  borderB = ""
  borderl = ""


   XWHERE = " and USUARIO = '" & SESSION("usuario") & "' "
                 XWHERE = XWHERE & " AND CODOBJETO > 0 "

                 If wkaplic = "OR" Or wkaplic = "PL" Then

                     XWHERE = XWHERE & " AND TIPOBJETO = 'RSK' "

                 Else

                     XWHERE = XWHERE & " AND TIPOBJETO <> 'RSK' "

                 End If

                 XORDER = " group by codobjeto ORDER BY fecha desc "


                 StrSql = ""
                 StrSql = StrSql & "SELECT CODOBJETO, max(TIPOBJETO), max(finicio) as fecha  "
                 StrSql = StrSql & " FROM OBJALOGDET B1"
                 StrSql = StrSql & " WHERE TACC = 'CONSEN'  " & XWHERE

                StrSql = StrSql & XORDER


set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

 if Not rs.EOF then rs.move(moverreg)

   Celda   = ""



    Fila=""
    POS=0
      xTrColor="#FFFFFF"
    do WHILE (NOT  rs.eof) AND (POS < tope)
       sxlink  = "ObjTreeSubTemaCons.asp?"
       sHref=""



         sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rs(0) & "&Tipo=" & rs(1) & "&MAINAPL=" & xMAINAPL
		                             sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
		                             sHref = "&nbsp;"

		                             sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


		                             sTipobjeto = EsqnomP(rs(1), rs(0), ofv)






		                             wmsgx = sTipobjeto & ""
		                             wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



		                             sImagen = wNomL

		                             If xarc = "NO" Then
		                                 xarc = "&nbsp;"
		                             End If

		                             sLink = sImagen

                            XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, sLink, "si")


          celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "' xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","",xarc  & XTSLNKM,"si")

         Fila    = Fila & ofv.GenRow("","aaa HEIGHT='25'","","","30","","",Celda,"si")

         If xTrColor = "#FFFFFF" Then
		        xTrColor = "#eeeeff"
		       Else
		        xTrColor = "#FFFFFF"
       End If
      RS.MOVENEXT
      POS=POS+1
    loop













   sHTML = sHTML & ofv.GenTabla("","aaa   ","","95% ","","0","","0","0","","",Fila,"si")

   if rs.eof then
   sxlinksig =""
   else
   sxlinksig =  "<A id=lnksig class=nt style='display:Xnone;'  HREF='BSQUltCons.asp?CANTREGMOVER=" & moverreg + tope  & "&tope="& tope &"&MAINAPL=" & MAINAPL &"'"
   sxlinksig = sxlinksig & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#052;</font></a>"
   end if

   if moverreg - tope< 0 then
   sxlinkant=""
   else
    sxlinkant =  "<A id=lnkant class=nt style='display:Xnone;' HREF='BSQUltCons.asp?CANTREGMOVER=" &  moverreg - tope  & "&tope="& tope  &"&MAINAPL=" & MAINAPL &"'"

	sxlinkant = sxlinkant & " TARGET=fesquemas><font size=3 color=red face='Corflaticon'>&#051;</font></a>"
   end if

	botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

  botones  ="<font size=2 color=red face=tahoma>Ver Más</font>" & botones
    sHTML = sHTML & botones



'''''''''''''

    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)
end if

response.write sHTML & "</body></html>"

Function EsqnomP(xtipo, xcodigo, ofv)

    Set xcn = ofv.conectar(ofv.strconn2)


    If xtipo = "RSK" Then

        StrSql = "SELECT * FROM  ORRSKCIRCUITO "
        StrSql = StrSql & " Where OBJID = 0" & xcodigo
        Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
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

   '''''sTipobjetoX = Capitalize(xdescri) & ": " & Capitalize(xobjdes)
        sTipobjetoX = xdescri & ": " & xobjdes

    Else


        sTipobjetoX = xcodigo

        StrSql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
        Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
        If Not RSM.EOF Then




            If SESSION("APLMDCLI") = "BC" Then
            sTipobjeto1 = (RSM(1)) & ":&nbsp;"
            sTipobjeto2 = RSM(3) & ""
            sTipobjeto3 = " - " & (RSM(2) & "")
            sTipobjeto4 = " (" & (RSM(5) & "") & " "
            sTipobjeto5 = (RSM(6) & "") & ") "
            Else
        '''''sTipobjeto1 = Capitalize(RSM(1)) & ":&nbsp;"
             sTipobjeto1 = RSM(1) & ":&nbsp;"
        '''''sTipobjeto2 = Capitalize(RSM(3) & "")
             sTipobjeto2 = RSM(3) & ""

       '''''sTipobjeto3 = " - " & Capitalize(RSM(2) & "")
            sTipobjeto3 = " - " & RSM(2) & ""
       '''''sTipobjeto4 = " (" & Capitalize(RSM(5) & "")
            sTipobjeto4 = " (" & RSM(5) & " )"
            If UCase(RSM(5) & "") = "O" Then
             sTipobjeto5 = ")"
            Else
       '''''sTipobjeto5 =  " " & Capitalize(RSM(6) & "") & ") "
            sTipobjeto5 = " " & RSM(6) & "" & ") "
            End If
            End If
            If SESSION("FULLDOCN") = "1" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2
            ElseIf SESSION("FULLDOCN") = "2" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
            Else
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
            End If



        End If

    End If





    Call ofv.cerrarconn(xcn)


    EsqnomP = sTipobjetoX


End Function



%>
