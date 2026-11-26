<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else



    Dim bconsultar()


    Generar = ""

    Addqry = ""
    Response.expires = 0

    Server.scripttimeout = 20
 ''   Response.buffer = True
    smenu = Request.querystring("Menu")
    stipoC = Request.querystring("tipoC")
    SEXPORT=REQUEST.FORM("EXPORT")

    SHTMLEXPORT="<TABLE>"


    if stipoC <> "" then
       SESSION("TIPOC") = stipoC
    end if

    sorden = Request.querystring("sorden")

    If Len(sorden) = 0 Then
      sorden = "N"
    End If

    Sfiltro = Request.querystring("sfiltro")

    If Len(Sfiltro) = 0 Then
      Sfiltro = "TODOS"
    End If

    Sfiltro = Replace(Sfiltro, "_", " ")

    strsql = ""

 'RESPONSE.WRITE SESSION("TIPOC") & "<BR>"

    sImagen1 = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
    sImagen2 = "<B><i class='las la-user';></i></B>"
    sImagen3 = "<B><i class='las la-angle-right' STYLE=COLOR:" & Session("CMFPNTIT") & ";></i></B>"
    sImagen4 = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>r</FONT></B>"


                 XFDES = REQUEST.FORM("DESDE3")  + clng(cdate("01-01-1900"))
                 XFHAS = REQUEST.FORM("HASTA3")  + clng(cdate("01-01-1900"))

                 XGDESDE = XFDES
                 XGHASTA = XFHAS
                 'response.write "XGDESDE: " & XGDESDE & "<br>" & "XGHASTA: " & XGHASTA & "<br>"

                 'strsql = strsql & " SELECT A.CODOBJETO, A.TIPOBJETO, CONVERT(VARCHAR(50),E.GRPUSR) AS GRPUSR,"
                 strsql = strsql & " SELECT A.CODOBJETO, A.TIPOBJETO, E.GRPUSR AS GRPUSR,"
                 strsql = strsql & " E.SECTOR, AVG(ISNULL(PROM,0)) AS PROMF, "
                 strsql = strsql & " SUM(ISNULL(CANT,0)) AS CANTF, MAX(ESTADO) AS ESTADO , MAX(VERSION) AS VERSION, "
                 strsql = strsql & " MAX(ISNULL(NROVER,0)) AS NROVER, MAX(FECVIGEN) AS FECVIGEN, "
                 strsql = strsql & " MAX(CODIDEN) AS CODIDEN, MAX(DESCRI) AS DESCRI, ISNULL('1','TODOS') AS USRDIST,"
                 strsql = strsql & " MAX(ISNULL(NOMUSR,'TODOS')) AS USRDISTDESC,"
                 strsql = strsql & " ISNULL(B.USUARIO,'***SIN') AS USRACC, MAX(ISNULL(B.GSECTOR,'0')) AS GSECTOR,"
                 strsql = strsql & " MAX(ISNULL('3','***SINGRP')) AS GGRUPO, MIN(ISNULL(CANT,0)) AS LEYO, "
                 strsql = strsql & " MAX(ISNULL(B.NOMUSR,'**SINDSC')) AS GNOMUSR ,max(d.nroacta) as nroacta"

                 strsql = strsql & " FROM OBJETOS A "

                 strsql = strsql & " INNER JOIN OBJPRTY D ON A.CODOBJETO = D.CODOBJETO "

                 strsql = strsql & " INNER JOIN ( SELECT F1.CODOBJETO, F1.TIPOBJETO,ISNULL(F2.GRPUSR,'TODOS') AS GRPUSR,"
                 strsql = strsql & " ISNULL(F2.SECTOR,0) AS SECTOR FROM OBJETOS F1 "
                 strsql = strsql & " LEFT JOIN OBJDIST F2 ON F1.CODOBJETO = F2.CODOBJETO"
                 strsql = strsql & " ) E ON A.CODOBJETO = E.CODOBJETO "

                 'strsql = strsql & " LEFT JOIN USUARIOS H "
                 'strsql = strsql & " ON ( (E.GRPUSR = H.ID) OR ((E.GRPUSR = 'SECT-') AND (E.SECTOR = H.SECTOR)) )"

                 strsql = strsql & " LEFT JOIN "

			if session("UN") = 1 then
                 strsql = strsql & " (select CODOBJETO, TIPOBJETO, USUARIO, AVG(HFINAL - HINICIO) AS PROM, "
                 strsql = strsql & " SUM(1) AS CANT, MAX(G2.SECTOR) AS GSECTOR, MAX(G2.GRUPO) AS GGRUPO, "
                 strsql = strsql & " MAX(G2.descripcion) AS NOMUSR "
                 strsql = strsql & " from objalogdet G1, USUARIOS G2 where tacc = 'CONSEN' AND G1.USUARIO = G2.ID "
                 strsql = strsql & " AND (G1.FINICIO > " & XGDESDE - 1 & " AND G1.FINICIO < " & XGHASTA + 1 & ")"
                 strsql = strsql & " GROUP BY CODOBJETO,TIPOBJETO,USUARIO)"
			else
				 strsql = strsql & " (select G1.CODOBJETO, G1.TIPOBJETO, G1.USUARIO, AVG(G1.HFINAL - G1.HINICIO) AS PROM, SUM(1) AS CANT, MAX(G2.SECTOR) AS GSECTOR,  "
				 strsql = strsql & " MAX(G2.usrname) AS NOMUSR "
				 strsql = strsql & " from objalogdet G1 "
				 strsql = strsql & " inner join objalog g2 on (g2.tipobjeto='SKRLOG' and g2.usuario=g1.usuario and g2.codemp=" & session("UN") & ") "
				 strsql = strsql & " where G1.tacc = 'CONSEN' AND (G1.FINICIO > " & XGDESDE - 1 & " AND G1.FINICIO < " & XGHASTA + 1 & ")"
				 strsql = strsql & " GROUP BY G1.CODOBJETO,G1.TIPOBJETO,G1.USUARIO )"
			end if



                 strsql = strsql & " B ON B.CODOBJETO = E.CODOBJETO AND B.TIPOBJETO = E.TIPOBJETO "
                 strsql = strsql & "    AND ( (E.GRPUSR = 'TODOS') OR (E.GRPUSR = B.USUARIO) OR "
                 strsql = strsql & " ((E.GRPUSR = 'SECT-') AND (E.SECTOR = B.GSECTOR)) )"

                 strsql = strsql & " WHERE A.CODOBJETO > 0 AND D.ESTADO = 3 "
                 'strsql = strsql & " AND (D.FECVIGEN > " & XGDESDE - 1 & " AND D.FECVIGEN < " & XGHASTA + 1 & ")"
              if session("UN") = 1 then
				 STRSQL = STRSQL & " AND B.USUARIO IS NOT NULL "
			  else
				 STRSQL = STRSQL & " AND B.USUARIO IS NOT NULL AND A.OBJUN=" & session("UN")
		      end if
                 strsql = strsql & " GROUP BY A.CODOBJETO,A.TIPOBJETO, E.GRPUSR, E.SECTOR, B.USUARIO"
                ' strsql = strsql & " ORDER BY D.FECVIGEN DESC,A.CODOBJETO DESC, LEYO DESC, E.SECTOR, USRACC, CANTF DESC"
                 strsql = strsql & " ORDER BY FECVIGEN DESC, A.CODOBJETO DESC, LEYO DESC, E.SECTOR, USRACC, CANTF DESC"

				'RESPONSE.Write(CHR(13) & CHR(13))
               'RESPONSE.WRITE STRSQL & CHR(13)
				'RESPONSE.Write(CHR(13) & CHR(13))

    CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
    CantRegMover = CDbl(Request.querystring("CantRegMover"))

        sverestado = " and objnivel >= " & Session("SegNivel") & " And Estcod In (9,3)"

        If Request.Form.Count > 0 Then

           sID = Request.Form("Desde3")
           sID2 = Request.Form("Hasta3")
           sID3 = Request.Form("Fecha")
           Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
        Else
           sID = Request.querystring("Desde")
           sID2 = Request.querystring("Hasta")
           sID3 = Request.querystring("Fecha")
           Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
        End If

        sMenuAdd = ""
        sMenuMas = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
        sMenuMas = sMenuMas & sID2 & "&Fecha=" & sID3


        xsid = OFV.convertircar(sID, "/", "")
        xsid = OFV.convertircar(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID = "") Then sID = "1-1-1950"
        xsid = OFV.convertircar(sID2, "/", "")
        xsid = OFV.convertircar(xsid, "-", "")
        If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"


    Addqry = "&Menu=" & smenu & Addqry
    starget = "estactalis.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry  &"&perfil="&replace(sperfil," ","_")
    sVolver = "estactalis.asp?Menu=" & smenu & sMenuAdd &"&perfil="&replace(sperfil," ","_")
    Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")
    Set cn = OFV.conectar(OFV.strconn2)

    sql = "select cliente from clienteplan where id=" & session("UN")
	set rx = ofv.crearconsultaEx(sql,cn,1,parametros)
	empresa = ""
	if not rx.eof then empresa = rx(0)
	Call OFV.cerrarconsulta(rx)

    sTitulo = "Esquemas Vigentes&nbsp;&nbsp;&nbsp;&nbsp;Desde:" & Fechador(sID) & "&nbsp;&nbsp;&nbsp;&nbsp;Hasta: " & Fechador(sID2) & "&nbsp;&nbsp;&nbsp;&nbsp;Empresa: " & empresa

  'response.write "sql: " & strsql & "<br>"
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

    sHTML = OFV.FormHeader(sTitulo)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs

    sHTML = sHTML & "<center>"

    Celda = OFV.GenCelda("", "ttt colspan=7", "", "", "", "", "", sTitulo, "si")

    Fila = OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")


    Celda = OFV.GenCelda("", "ttr", "", "5%", "", "", "", "*", "si")
    Celda = Celda & OFV.GenCelda("", "ttr colspan=3", "", "50%", "", "", "", "ESQUEMA NORMATIVO", "si")
    Celda = Celda & OFV.GenCelda("", "ttr", "", "15%", "", "", "", "TIPO", "si")
    Celda = Celda & OFV.GenCelda("", "ttr COLSPAN=2 ", "", "30%", "", "", "", "VIGENCIA", "si")


    Fila = fila & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")

    IF SEXPORT = "XLS" THEN
       SHTMLEXPORT=SHTMLEXPORT & "<TR><TD>" & "*" & "</TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "SITE/SUCURSAL" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "SECTOR" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "PUESTO/PERFIL" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "LEGAJO" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "USUARIOS QUE ACCEDIERON" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "TIPO" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "ID ESQUEMA" & "</U></TD>"
       'mauro 09/05/17
       'SHTMLEXPORT=SHTMLEXPORT & "<TD></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "ESQUEMA NORMATIVO" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "NRO DE ACTA" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "PROMEDIO" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "FECHA DESDE" & "</U></TD>"
       SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "FECHA HASTA" & "</U></TD></TR>"
    END IF

    nLinea = 0
   ' response.write
    fi = CLng(CDate("01-01-1900"))
     PosReg = 0
     PosReg2 = 0
     If Not rs.EOF THEN

        XSANTOBJ  = clng(RS(0))
        XSANTGRPT = "NO"
        XSANTLEIN = "NO"
' RESPONSE.WRITE XSANTOBJ  & "<BR>"
     END IF
sss = 0
saa = 0
saav     = "S"

     Do While (Not rs.EOF)

         if saa <> rs(0) then   saav     = "S"

          ' RESPONSE.WRITE XSANTOBJ  & " - LOOP1<BR>"

         'sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"
		 sImagen = "<i class='las la-history' ></i>" 'color:#993333

         sss = sss + 1

         dVigencia = "SD"
         If Not IsNull(rs(9)) Then dVigencia =fechadorcor( CDate(CLng(rs(9)) + clng(cdate("01-01-1900"))),"") 'CDate(CLng(rs(9))) + clng(cdate("01-01-1900"))
         dVign = InStr(dVigencia, "1900")
         If dVign > 0 Then dVigencia = "SD"

         sTipobjeto = RS(1)
         sCODobjeto = RS(0)
         sobjeto    = RS(0)

         Strsql = "Select * From Tipobjeto Where tipobjeto='" & sTipobjeto & "'"

		 'RESPONSE.Write(CHR(13) & CHR(13))
         'RESPONSE.WRITE STRSQL & CHR(13)
		 'RESPONSE.Write(CHR(13) & CHR(13))

         set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         IF NOT RSX.EOF THEN sTipobjeto = RSX(1)


          sobjeto = RS(11) & "-" & RS(10)

        IF  RS(7) = "O" THEN
            sobjeto = sobjeto & "&nbsp;(O)"
        ELSE
            sobjeto = sobjeto & "&nbsp;(V&nbsp;" & RS(8) & ")"
        END IF


        sActcon  = "CoESTDLISDET.asp?Codigo=" & rs("codobjeto") & "&Tipo=" & rs("tipobjeto")
        sActcon  = sActcon & "&FDES=" & XFDES & "&FHAS=" & XFHAS
        sHref    = "&nbsp;"
        sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)

        Celda    = ""
        sLink    = OFV.GenLink("", "istt ", sHref, "", "", sFuncion, sImagen, "si")
        Celda    = Celda & OFV.GenCelda("", "", "center", "", "", "", "", sLink, "si")

        sLink    = OFV.GenLink("", "istt", sHref, "", "", sFuncion, sobjeto, "si")
        Celda    = Celda & OFV.GenCelda("", "ttt  COLSPAN=3 ", "", "", "", "", "", sLink, "si")

        Celda    = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "",sTipobjeto, "si")
        Celda    = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", dVigencia, "si")
        Celda    = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "","&nbsp;", "si")

         if saa <> rs(0) then

            Fila     = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

         else
            saav     = "N"

         end if

        XSANTLEIN = "X"
        FilaX = ""

     IF (RS(2) & "") = "TODOS" AND RS(17) = 0 THEN

         if saa <> rs(0) then

              Celda = ""
              Celda = Celda & OFV.GenCelda("", "UT COLSPAN=2 ", "", "", "", "", "", "&nbsp;", "si")
              Celda = Celda & OFV.GenCelda("", "UT  ", "RIGHT", "10%", "", "", "", sImagen4 & "&nbsp;", "si")
              Celda = Celda & OFV.GenCelda("", "UT COLSPAN=3 ", "", "", "", "", "", "<U>NO SE REGISTRAN ACCESOS PARA EL DOCUMENTO.!!</U>", "si")

              FilaX = FilaX & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")
              saa   = RS(0)

         end if

              rs.movenext

      ELSE


     Do While (Not rs.EOF)



        IF NOT  clng(XSANTOBJ) = clng(RS(0)) THEN EXIT DO

' RESPONSE.WRITE XSANTOBJ  & " - LOOP2<BR>"

           IF XSANTLEIN = "X" THEN
              XSANTLEIN = "S"
              XTABLAX = "USUARIOS QUE ACCEDIERON"
              XTABLAXT1 = "<SMALL>ACCESOS</SMALL>"
              XTABLAXT2 = "<SMALL>PROMEDIO</SMALL>"
           ELSE
              XSANTLEIN = "N"
              XTABLAX = "USUARIOS QUE NO ACCEDIERON"
              XTABLAXT1 = "&nbsp;"
              XTABLAXT2 = "&nbsp;"
           END IF

              Celda = OFV.GenCelda("", "VT", "RIGHT", "5%", "", "", "", sImagen3 & "&nbsp;", "si")
              Celda = Celda & ofv.GenCelda("","vt  colspan=3 ","","","","","", XTABLAX ,"si")
              Celda = Celda & ofv.GenCelda("","vt   ","RIGHT","","","","", XTABLAXT1 ,"si")
              Celda = Celda & ofv.GenCelda("","vt   ","RIGHT","","","","", XTABLAXT2 ,"si")
              FilaX = FilaX & ofv.GenRow("","","","","30","","",Celda,"si")

        Do While (Not rs.EOF)

        IF NOT  clng(XSANTOBJ) = clng(RS(0)) THEN EXIT DO

             IF XSANTLEIN = "S" THEN
                IF RS(17) = 0 THEN EXIT DO
             END IF



' RESPONSE.WRITE XSANTOBJ  & " - LOOP3<BR>"

           IF XSANTLEIN = "S" THEN
              XGCANT = RS(5)
              XGPROM = GETGLBHMS(RS(4)) & "&nbsp;"
           ELSE
              XGCANT = "&nbsp;"
              XGPROM = "&nbsp;"
           END IF

              XUSRACC = RS13 & "X"

              IF RS(13) = "TODOS" THEN
                 xxx=getInfoUserEX(RS(14),session("un"))
                 XGUSRDSC = session("SGU_nombre")
                 'XGUSRDSC = RS(18)

                 'IF XGUSRDSC = "**SINDSC" THEN XGUSRDSC = "* SIN USUARIOS ASIGNADOS"
                  IF XGUSRDSC = "NO" THEN XGUSRDSC = "* SIN USUARIOS ASIGNADOS"
              ELSE
                 XGUSRDSC = RS(13)
              END IF

              IF RS(17) = 0 THEN
                 XGUSRSET = sImagen1
              ELSE
                 XGUSRSET = sImagen2

                 sActcon = "CoESTDLISUSRDET.asp?USR=" & rs(14) & "&COD=SI&Codigo=" & rs(0) & "&Tipo=" & rs(1)
                 sActcon = sActcon & "&FDES=" & XFDES & "&FHAS=" & XFHAS
                 sHref = "&nbsp;"
                 sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)
                 sLink = OFV.GenLink("", "NT", sHref, "", "", sFuncion, XGUSRSET, "si")

                 XGUSRSET = sLink

                 sLink = OFV.GenLink("", "NT", sHref, "", "", sFuncion, XGUSRDSC, "si")

                 XGUSRDSC = sLink

              END IF

              Celda = ""
              Celda = Celda & OFV.GenCelda("", "UT COLSPAN=2 ", "", "", "", "", "", "&nbsp;", "si")
              Celda = Celda & OFV.GenCelda("", "UT  ", "RIGHT", "10%", "", "", "", XGUSRSET & "&nbsp;", "si")
              Celda = Celda & OFV.GenCelda("", "UT ", "", "", "40%", "", "", XGUSRDSC, "si")
              Celda = Celda & OFV.GenCelda("", "UT ", "RIGHT", "15%", "", "", "", XGCANT & "&nbsp;", "si")
              Celda = Celda & OFV.GenCelda("", "UT ", "RIGHT", "20%", "", "", "", XGPROM & "&nbsp;", "si")
              FilaX = FilaX & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")

              rs.movenext

             If rs.EOF THEN EXIT DO

          If Not rs.EOF AND clng(XSANTOBJ) = clng(RS(0)) THEN

             IF XSANTLEIN = "S" THEN
                IF RS(17) = 0 THEN EXIT DO
             END IF
          END IF
        LOOP

             IF SEXPORT = "XLS" THEN
                          if not rs.EOF then
		                  xxx=getInfoUserExBp(RS(14))

		                   XGUSRDSC = session("SGU_nombre")
		                   userid=  session("SGU_userid")
		                   CODPST=session("SGU_codigopuesto")
		                   SGU_legajo=session("SGU_legajo")
		                   SGU_site=session("SGU_site")
		                   SGU_siteDescri=session("SGU_siteDescri")
		                   else
		                   XGUSRDSC="&nbsp;"
		                   userid="&nbsp;"
		                   CODPST="&nbsp;"
		                   SGU_legajo="&nbsp;"
		                   SGU_site="&nbsp;"
		                   SGU_siteDescri="&nbsp;"

		                  end if

		                  strsql="select descri from tipoperfil where tipoperfil='" & CODPST & "'"
		                '  response.write strsql
		                  codpst="&nbsp;"
		                  set RstPef=ofv.crearconsultaex(strsql, cn, 1, "")
		                  if not RstPef.eof then CODPST=RstPef(0)
		                  call ofv.cerrarconsulta(RstPef)


		                  CODinter="&nbsp;"
		                  ecodescri="&nbsp;"
		                  if not rs.EOF then
		                  strsql="select descri,codinter from eco where eco='" & rs(15) & "'"

						  set RstPef=ofv.crearconsultaex(strsql, cn, 1, "")
						  if not RstPef.eof then
						  CODinter=RstPef(1)
						  ecodescri=RstPef(0)
						  end if
						  else
						  CODinter="&nbsp;"
						  ecodescri="&nbsp;"
						  end if

		                  call ofv.cerrarconsulta(RstPef)
                         SHTMLEXPORT=SHTMLEXPORT & "<TR><TD>"&  SGU_site & "</TD>"
                         SHTMLEXPORT=SHTMLEXPORT & "<TD>"& SGU_siteDescri & "</TD>"
				         SHTMLEXPORT=SHTMLEXPORT & "<TD>"& ecodescri & "</TD>"
				         'SHTMLEXPORT=SHTMLEXPORT & "<TD>"& CODinter & "</TD>"
				         SHTMLEXPORT=SHTMLEXPORT & "<TD>" & CODPST & "</TD>"
				         SHTMLEXPORT=SHTMLEXPORT & "<TD>" & SGU_legajo & "</TD>"
				         SHTMLEXPORT=SHTMLEXPORT & "<TD>" & XGUSRDSC & "</TD>"

				        if not rs.eof then
				          SHTMLEXPORT=SHTMLEXPORT & "<TD>" & RS(1) & "</TD>"
				          SHTMLEXPORT=SHTMLEXPORT & "<TD>" & RS(0) & "</TD>"
				          SHTMLEXPORT=SHTMLEXPORT & "<TD>" & RS(11) & "</TD>"
				          SHTMLEXPORT=SHTMLEXPORT & "<TD>" & RS("nroacta") & "</TD>"
				        else
				          SHTMLEXPORT=SHTMLEXPORT & "<TD></TD>"
						  SHTMLEXPORT=SHTMLEXPORT & "<TD></TD>"
				          SHTMLEXPORT=SHTMLEXPORT & "<TD></TD>"
				          SHTMLEXPORT=SHTMLEXPORT & "<TD></TD>"
				        end if


				         SHTMLEXPORT=SHTMLEXPORT & "<TD>" & XGPROM & "</TD>"
				         SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDATE(XFDES),"") & "</TD>"
						 SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDATE(XFHAS),"") & "</TD></TR>"
        END IF

     LOOP
    END IF

              Celda = OFV.GenCelda("", "VT COLSPAN=6 ", "", "", "", "", "", "&nbsp;", "si")
              FilaX = FilaX & ofv.GenRow("","","","","","","",Celda,"si")

    XTABLAX = OFV.GenTabla("", "", "", "100%  ", "", "0", "", "-1", "-2", "", "", FilaX, "si")

  if saav <> "N" then
    Celda = ofv.GenCelda("","ut","","","","","", "&nbsp;" ,"si")
    Celda = Celda & ofv.GenCelda("","uti  colspan=6 style={border-style:solid;border-width:1;border-color:#dddddd;} ","","","","","", XTABLAX ,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    PosReg2 = PosReg2 + 1

  end if

     If Not rs.EOF THEN

        XSANTOBJ  = clng(RS(0))
        XSANTGRPT = "NO"
        XSANTLEIN = "N"
' RESPONSE.WRITE XSANTOBJ  & " - CHECK2<BR>"
     END IF

        PosReg = PosReg + 1

    Loop

    if PosReg > 0 then
       Celda = ofv.GenCelda("","ttt  colspan=7 ","","","","",""," TOTAL: " & posreg2 & " ESQUEMAS             " & sInput,"si")
    else
       Celda = ofv.GenCelda("","ttt  colspan=7 ","","","","",""," NO HAY ESQUEMAS PARA LA PRESENTE SELECCION ","si")
    end if

    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    XTABLA = OFV.GenTabla("", "aaa", "", "100%", "", "0", "", "-1", "0", "", "", Fila, "si")


    Formulario = XTABLA

    sHTML = sHTML & Formulario

     sAccion = "COESTDSEGUI.asp"
    Celda   = ofv.BotonVolver(sAccion)

    sHTML = sHTML & OFV.botoneraESSX("0", "", "", "", "", "", "", "", "", "", "", "", Celda)

    sHTML = sHTML & "</body></html>"

     session("sperfil")=""


    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

''     Generar = sHTML

IF SEXPORT ="XLS" THEN
 Response.ContentType = "application/vnd.ms-excel"
 Response.AddHeader "content-disposition", " filename=CoEstdLisSEG.xls"
 response.write SHTMLEXPORT & "</TABLE>"
ELSE
 response.write sHTML
' response.flush
END IF

end if


'response.write sHTML
'response.flush

Function Fechador(wkfecha)

      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)

End Function

%>