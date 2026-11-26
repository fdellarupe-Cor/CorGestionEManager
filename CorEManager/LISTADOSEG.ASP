<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

     Generar = ""

     Addqry = ""
     Response.expires = 0

     Server.scripttimeout = 600
     Response.buffer = True


     'Response.WRITE request.querystring("MODO")
     '

     XXXX    = NEWEMPRE("NO")
     smenu   = Request.querystring("MENU")
     sorden  = Request.querystring("sorden")
     xTipoObj  = Request.querystring("xtipo")


     SMENU = TRIM(sMenu)

	base=request.querystring("BASE")

      IF base <> "" THEN
         SESSION("BASE")=base
      END IF


  if Session("CorApli") = "COREM20" then
     wkaplic = "EM"
  elseif Session("CorApli") = "COREP20" then
     wkaplic = "PL"
  else
     wkaplic = "OR"
  end if


     If Len(sorden) = 0 Then
         sorden = "N"
     End If

     Sfiltro = Request.form("sfiltro")

        If Len(Sfiltro) = 0 Then
           Sfiltro = Request.querystring("sfiltro")
        End If

     If Len(Sfiltro) = 0 Then
         Sfiltro = "TODOS"
     End If

     Sfiltro = Replace(Sfiltro, "_", " ")
     CantRegMostrar = 10
     CantRegMover = CDbl("0" &Request.querystring("CantRegMover"))
     ReDim Preserve bconsultar(4)
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "4"
     bconsultar(4) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "2"
     bconsultar(2) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "1"
     bconsultar(1) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "3"
     bconsultar(3) = OFV.AConsultar

	 'response.write Session("Form") & bconsultar(3)
     For i = 1 To 4
         If bconsultar(i) Then
            sverestado = sverestado & i & ","
            if i = 1 then sverestado = sverestado & "6" & ","
            if i = 2 then sverestado = sverestado & "5" & ","

         end if
     Next
     sverestado = "(" & Left(sverestado, Len(sverestado) - 1) & ") "


      addfrom =  ""


        If smenu = "CONSGRAL" Then





            FESTADO         = cstr(request.form("estado") & "")

            fMODO          = cstr(request.form("MODO")& "")

            FDESDE          = cstr(request.form("DESDE3")& "")
            FHASTA          = cstr(request.form("HASTA3")& "")



            if FESTADO <> ""  then
            session("FESTADO")=FESTADO
            else
            FESTADO=session("FESTADO")
            end if




             if FDESDE <> ""  then
			            session("FDESDE")=FDESDE
			            else
			            FDESDE=session("FDESDE")
            end if
             if FHASTA <> ""  then
			            session("FHASTA")=FHASTA
			            else
			            FHASTA=session("FHASTA")
            end if
			if FMODO <> ""  then
	 			            session("FMODO")=FMODO
	 			            else
	 			            FMODO=session("FMODO")
            end if

			if FDESDE & "" = "" then FDESDE = request.querystring("desde3")
			if FHASTA & "" = "" then FHASTA = request.querystring("hasta3")
			if FMODO & "" = "" then FMODO = request.querystring("modo")
			if FESTADO & "" = "" then FESTADO = request.querystring("estado")

			'response.write "FMODO: " & FMODO




            'response.write  request.form

        End If
           scal=request.querystring("scal")
 If smenu = "PORACTA" Then



  ElseIf smenu = "PORIDENTIFICACION" Then


  Else

        If  Robj = "F" Then

            SPerfil = Request.Form("perfil") & ""
            sID     = Request.Form("lista") & ""
            sID2    = Request.Form("lista") & ""
            sID3    = Request.Form("buscar") & ""

            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil

            SPerfil = Replace(SPerfil, "_", " ")

        Else

            SPerfil = Request.querystring("perfil") & ""
            sID     = Request.querystring("Lista") & ""
            sID2    = Request.querystring("Lista") & ""
            sID3    = Request.querystring("buscar") & ""
            Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
            SPerfil = Replace(SPerfil, "_", " ")

        End If

        sID         = OFV.convertircar(sID, "_", " ")
        If sID = "" Then sID = sID3

        sMenuAdd    = ""
        sMenuMas    = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
        sMenuMas    = sMenuMas & sID3 & "&perfil=" & SPerfil

    End If



    Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre

    starget = "RESPUESTA2.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry &"&DESDE=" & FDESDE & "&HASTA=" & FHASTA
    sVolver = "RESPUESTA2.asp?Menu=" & smenu & sMenuAdd

    Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")

  '  response.write sorden &" - "

    'response.write Addqry &

     Select Case sorden
                Case "N"
                        sTitOrd = sTitOrd & "Nombre"
                        strsql  = strsql & " order by a.descri,csporfechas.fecvigen "
                Case "B"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by C.descri,csporfechas.fecvigen "
                Case "P"
                        sTitOrd = sTitOrd & "Producto"
                       strsql  = strsql & " order by  tp.descri,d.fecvigen desc "
                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by  EST.estado, A.DESCRI "
                Case "C"
                        sTitOrd = sTitOrd & "Codigo"
                        strsql  = strsql & " order by  b.codigoform "
                Case "V"
                        sTitOrd = sTitOrd & "Vigencia"
                        strsql  = strsql & " order by  csporfechas.fecvigen desc "
                Case "A"
                        sTitOrd = sTitOrd & "Autor"
                        strsql  = strsql & " order by  d.autor "
        End Select


          contar=""

    Set cn = OFV.conectar(OFV.strconn2)
     Set cn5 = OFV.conectar(OFV.strconn5)
    cn.CommandTimeout = 0



            newstrsql=""


    SELECT CASE FMODO

    CASE "SG"

            SQL= " select a.fechafin,a.objid,a.nombre,a.solicitante,a.responsable,a.estado,b.estado AS EST_DESCRI,tipobjeto from cs_pedidos a "

            if session("BASE") = "K" then
            	sql=sql & " inner join estadosform b on (b.objid=a.estado and b.tipo='PD') "
            else
            	sql=sql & " inner join estadosForm b on (b.objid=a.estado and b.tipo='SN') "
			end if

            NEWSTRSQL= ""
            Sand=""
            if Festado  <> "" AND Festado <>"0" then

              NEWSTRSQL= " A.estado='"& Festado & "'"
              Sand= " and "
            end if

            if (FDESDE  <> "" AND FDESDE <>"0" ) and (FHASTA  <> "" AND FHASTA <>"0" )  then

             NEWSTRSQL= NEWSTRSQL & Sand &  "   a.fechafin  > = '"& FDESDE &"' AND a.fechafin <= '"& FHASTA &"' "
              Sand= " and "
            end if


           if   NEWSTRSQL & "" <> "" then

           		if session("BASE") = "K" then
           			SQL=sql & " where " &  NEWSTRSQL & " and a.tipobjeto='PDCORXY'"
            	else
					SQL=sql & " where " &  NEWSTRSQL & " and a.tipobjeto='SN'"
            	end if

		   else

				if session("BASE") = "K" then
					SQL=sql & " where a.tipobjeto='PDCORXY'"
			   	else
					SQL=sql & " where a.tipobjeto='SN'"
            	end if

           end if



   CASE "RANK"

            SQL= "select  SUM(1) "
			sql=sql & "  from cs_pedidos a "

			if session("BASE") = "K" then
			    sql=sql & " inner join estadosform b on (b.objid=a.estado and b.tipo='PD') "
			else
				sql=sql & " inner join estadosForm b on (b.objid=a.estado and b.tipo='SN') "
			end if

			if session("BASE") = "K" then
            	sql=sql & "  where a.tipobjeto='PDCORXY' and a.fechafin  > = '"& FDESDE &"' AND a.fechafin <= '"& FHASTA &"' "
			else
				sql=sql & "  where a.tipobjeto='SN' and a.fechafin  > = '"& FDESDE &"' AND a.fechafin <= '"& FHASTA &"' "
			end if


            Set rs = OFV.crearconsultaEx(sql, cn, 1, "")

            IF NOT RS.EOF THEN
             TOTAL=RS(0)
            ELSE
             TOTAL=0
            END IF
            CALL OFV.CERRARCONSULTA(RS)


            SQL= " SELECT *,SUM(1) AS TOTAL FROM ("
			sql=sql & "  select   B.ESTADO AS ESTADODSC,A.ESTADO AS ESTADO "
			sql=sql & "  from cs_pedidos a   "

			if session("BASE") = "K" then
				sql=sql & " inner join estadosform b on (b.objid=a.estado and b.tipo='PD') "
			else
				sql=sql & " inner join estadosForm b on (b.objid=a.estado and b.tipo='SN') "
			end if

			 if (FDESDE  <> "" AND FDESDE <>"0" ) and (FHASTA  <> "" AND FHASTA <>"0" )  then

			 	if session("BASE") = "K" then
			 		sql=sql & " WHERE a.tipobjeto='PDCORXY' and a.fechafin  > = '"& FDESDE &"' AND a.fechafin <= '"& FHASTA & "') T "
				else
					sql=sql & " WHERE a.tipobjeto='SN' and a.fechafin  > = '"& FDESDE &"' AND a.fechafin <= '"& FHASTA & "') T "
				end if


			END IF

            sql=sql & "  GROUP BY T.ESTADODSC,T.ESTADO "
            sql=sql & " order by total desc"

   END SELECT

	'response.write sql
    Set rs = OFV.crearconsultaEx(sql, cn, 1, parametros)



    sHTML = ofv.FormHeader(session("FormN")) & chr(13)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " var h=570;w=790 ;"
	if  session("APLMDCLI")="BC" THEN
      xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	END IF

    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & ","' & Chr(34)
    'xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
    xs = xs &"  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'" '& CHR(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs


 '   response.write Request.querystring & "<br>"
 '   response.write Request.form & "<br>"

  sImagenM = "<B><FONT FACE=""WEBDINGS"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""WEBDINGS"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI = "<B><FONT FACE=""WINGDINGS"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""WEBDINGS"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=WingDINGS COLOR=#aa0000 SIZE=4>x</FONT></B>"
  sImagena = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=2>4</FONT></B>"



  Fila    = ""
  Celda   = ""
  sHref   = ""


  Fila = ""



       FDESDE= dateserial(year(CDate(FDESDE)),month(CDate(FDESDE)),day(CDate(FDESDE))+2)
       FHASTA= dateserial(year(CDate(FHASTA)),month(CDate(FHASTA)),day(CDate(FHASTA))+2)

	   'response.write FDESDE & "--" & FHASTA

       sTitulo= "DESDE: " & fechatostr(FDESDE) & "&nbsp;&nbsp;-&nbsp;&nbsp;HASTA: " & fechatostr(FHASTA)



    nLinea = 0

    fi = CLng(CDate("01-01-1900"))
    PosReg = 0

    xTrColor = "#FFFFFF"

	if session("BASE") <> "K" then
		Celda = OFV.GenCelda("", "ttx colspan=7","", "10%", "","","",sTitulo, "si")
	else
		Celda = OFV.GenCelda("", "ttx colspan=6","", "10%", "","","",sTitulo, "si")
	end if
	Fila =  OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

    IF  FMODO="SG" THEN
      Celda = OFV.GenCelda("", "ttx align=center", "", "10%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=N target=_self ><B><u>&nbsp;FECHA&nbsp;</u></B></A>", "si")
	  Celda = CELDA & OFV.GenCelda("", "ttx align=center", "", "5%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=B target=_self ><B><u>&nbsp;Nº SOLICITUD&nbsp;</u></B></A>", "si")
	  if session("BASE") = "K" then
	  		Celda = CELDA &  OFV.GenCelda("", "ttx align=center", "", "30%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=P target=_self ><B><u>&nbsp;DESCRIPCION&nbsp;FORMULARIO&nbsp;</u></B></A>", "si")
      else
			Celda = CELDA &  OFV.GenCelda("", "ttx align=center", "", "30%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=P target=_self ><B><u>&nbsp;DESCRIPCION&nbsp;NORMATIVA&nbsp;</u></B></A>", "si")
      end if
      Celda =  CELDA &  OFV.GenCelda("", "ttx align=center", "", "10%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=C target=_self ><B><u>&nbsp;ORIGEN&nbsp;</u></B></A>", "si")
      Celda =  CELDA &  OFV.GenCelda("", "ttx align=center", "", "25%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=C target=_self ><B><u>&nbsp;DESTINO&nbsp;</u></B></A>", "si")
      Celda =  CELDA &  OFV.GenCelda("", "ttx align=center", "", "10%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=C target=_self ><B><u>&nbsp;ESTADO&nbsp;</u></B></A>", "si")
     ELSE
      Celda = OFV.GenCelda("", "ttx colspan=1", "", "40%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESTADO&nbsp;</u></B></A>", "si")
	  Celda = CELDA & OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=B target=_self ><B><u>&nbsp;RANKING&nbsp;</u></B></A>", "si")
    END IF

	 if session("BASE") <> "K" and FMODO = "SG" then
	 	Celda = CELDA & OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt Xhref=" & starget & "&sorden=B target=_self ><B><u>&nbsp;INFO&nbsp;</u></B></A>", "si")
	 end if
     Fila = Fila &  OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")
    If Not rs.EOF Then
     rs.Move (CantRegMover)

     POSREG=0
     aColores =  Array("red","orange","green","blue","brown","purple","gray")
      colorIndex = 0
     Do While (Not rs.EOF) And (PosReg < CantRegMostrar)
        Celda=""



      SELECT CASE FMODO

      CASE "SG"
        Celda = Celda & OFV.GenCelda("", "ut align=center", "", "10%", "", "", "", fechatostr(fechaConvert(RS("FECHAFIN") + fi ))  , "si")
        Celda = Celda & OFV.GenCelda("", "ut align=center", "", "10%", "", "", "",  RS("OBJID") &"&nbsp;&nbsp;&nbsp;&nbsp;", "si")
        Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "", RS("NOMBRE"), "si")

        sql="select b.codigo from  segcircuito a "

        SQL=SQL & " inner join segcktresp b on (b.segckt=a.segckt and modo='L' and orden <> 20) "

         SQL=SQL & "  where a.objcod='" & RS("OBJID")& "' and objtipo='" & RS("TIPOBJETO")& "' "

        SQL=SQL & "  ORDER BY B.ORDEN"
      '  response.write sql & "<br>/////<br>"

	         sql=""


       SQL=SQL & " select c.RESPASG, c.id, c.pscdsc "
       SQL=SQL & "   from  segcircuito a  "
       SQL=SQL & " inner join segpaso c on (c.segckt=a.segckt) "
       SQL=SQL & " where a.objcod='" & RS("OBJID")& "' and objtipo='" & RS("TIPOBJETO")& "'   "

       SQL=SQL & " ORDER BY c.ID desc "


         Set rsaut = OFV.crearconsultaEx(sql, cn5, 1, "")
         delim=""
         srespo=""
      '   Do While (Not rsaut.EOF)

	     if (Not rsaut.EOF) then

          responsable =rsaut(0)
        '  response.write responsable
          XX=GETINFOUSER(responsable)
          srespo= srespo & delim & SESSION("SGU_NOMBRE") & " /<br>" & rsaut(2)
          delim="<br>"

		  end if

       '   rsaut.movenext
       '  loop

         call ofv.cerrarconsulta(rsaut)


         XX=GETINFOUSER(RS("solicitante"))

        Celda = Celda & OFV.GenCelda("", "ut align=center", "", "10%", "", "", "",  SESSION("SGU_NOMBRE"), "si")
       '  XX=GETINFOUSER(RS("responsable"))
        Celda = Celda & OFV.GenCelda("", "ut align=center", "", "10%", "", "", "",   srespo, "si")
        Celda = Celda & OFV.GenCelda("", "ut align=center", "", "10%", "", "", "",  RS("EST_DESCRI"), "si")

        if session("BASE") <> "K" and (rs(5) = 4 or rs(5) = 5)  then
        	sImagenP = "&#0232;"
			sinput = "<a class='ft1n' href='listadosegdet.asp?cod=" & rs(1) & "&tipo=" & rs(7) & "&" & request.form & "' title='Ver desarrolladores' >" & sImagenP & "</a>"
			Celda = Celda & OFV.GenCelda("", "ut ", "center", "10%", "", "", "",sinput, "si")
		else
			Celda = Celda & OFV.GenCelda("", "ut ", "center", "10%", "", "", "","&nbsp;", "si")
        end if


        PosReg = PosReg + 1
      CASE "RANK"
         ANCHO=400
         TOTALp= RS("TOTAL")
         PORC=(TOTALP * 100) / TOTAL
         LARGO= (PORC * ANCHO) /100
         LARGO =CINT(LARGO)
      '   RESPONSE.WRITE LARGO
         Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "", RS("ESTADOdsc"), "si")

         Celda = Celda & OFV.GenCelda("", "ut ", "", "90%", "", "", "", "<DIV STYLE='border-radius:5px;height:25px;background-color:"& acolores(colorIndex)& ";width:" & LARGO & "px;'   ID=DIVEST_"  & RS("ESTADO") & ">&nbsp;</DIV>" & TOTALp & " SOLICITUDES", "si")
         colorIndex = colorIndex + 1

         colorIndex=colorIndex mod ubound(aColores)
      END SELECT


        Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")


		 if xTrColor = "#FFFFFF" then
			xTrColor = "#F0F0F0"
		 ELSE
			xTrColor = "#FFFFFF"
		 END IF


        rs.movenext



    Loop

    if FMODO="RANK" THEN

	       'XFRASE = "<BR><BR><font color=#993333 SIZE=3 FACE='WEBDINGS' >i</font>"
		   XFRASE = "<BR><BR><i class='las la-info-circle' style=color:#993333;font-size:10pt;></i>"
	       XFRASE = XFRASE &  "&nbsp;TOTAL DE SOLICITUDES :" & TOTAL & "<BR><BR>"
	      Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", XFRASE , "si")
	       xFila =  OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")
	else
	xfila=""
    END IF

   else

        'XFRASE = "<BR><BR><font color=#993333 SIZE=3 FACE='WEBDINGS' >i</font>"
		XFRASE = "<BR><BR><i class='las la-info-circle' style=color:#993333;font-size:10pt;></i>"
        XFRASE = XFRASE &  "&nbsp;NO SE ENCONTRARON SOLICITUDES PARA LA BUSQUEDA SOLICITADA<BR><BR>"

        Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", XFRASE, "si")

        Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")


   end if

    fila=fila & xfila

    sHTML = sHTML & OFV.GenTabla("", "", "", "100% class=aaa", "", "0", "", "-1", "0", "", "", Fila, "si")
    'response.write Addqry


       sAccion = "SEGUIMIENTOSOLICITUD.asp"


    Celda   = ofv.BotonVolver(sAccion)


    sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)


    sHTML = sHTML  & "</body></html>"


    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)
     Call OFV.cerrarconn(cn5)


    Generar = sHTML

'''''''''''''''''''

  sHTML = Generar

 '' sHTML = oClc.Generar



end if

response.write sHTML
response.flush



Function fechaConvert(wkfecha)
  Nfecha=cdate(wkfecha)
  Nfecha=dateserial(year(wkfecha),month(wkfecha),day(wkfecha))
  fechaConvert=Nfecha
end function

Function fechatostr(wkfecha)
  fechatostr=cdate(wkfecha)
  fechatostr=day(wkfecha) & "/" & month(wkfecha) & "/" & year(wkfecha)

end function



%>