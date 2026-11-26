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

    Server.scripttimeout = 600
    Response.buffer = True

    XFNAME = request.querystring("XFNAME")
    IF XFNAME = "" THEN XFNAME = "NO"


    smenu = Request.querystring("Menu")

    sorden = Request.querystring("sorden")

    If Len(sorden) = 0 Then
      sorden = "N"
    End If

    Sfiltro = Request.querystring("sfiltro")

    If Len(Sfiltro) = 0 Then
      Sfiltro = "TODOS"
    End If

    Sfiltro = Replace(Sfiltro, "_", " ")

  '  RESPONSE.WRITE Request.querystring
    CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
    CantRegMover = CDbl(Request.querystring("CantRegMover"))

    selact=Request.querystring("selact")


        'sverestado = " and objnivel >= " & Session("SegNivel") & " And Estcod In (9,3)"

estado = request.querystring("estado")


           sID = Request.querystring("Desde")
           sID2 = Request.querystring("Hasta")
           sID3 = Request.querystring("Fecha")
           Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3


     xin = ""
     xinc = 0

        IF XFNAME = "NO" THEN
           for each x in request.form

              if request.form(x) = "1" then
                 xinc = xinc + 1
                 xin = xin & replace(x,"M","") & ","
'response.write x & ": " & request.form(x) & " <br>"
              end if
           next


'response.write xinc

           if xinc > 0 then

              xin = left(xin,len(xin)-1)

              if xinc = 1 then
                 xin = " and codobjeto = " & xin & " "
              else
			  'INC 1439: en el renglón de abajo se agregó la palabra not porque no traía los esquemas correctos. Se copió de Bcp Galicia.
                 xin = " and codobjeto   in (" & xin & ") "
              end if

            end if

            session("XIN") = xin

        ELSE
             xin = session("XIN")
        END IF






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


        strsql = "select * from csporfechas "
       ' cdesde = CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
       ' chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

        cdesde = clng(sID) - 1
        chasta = clng(sID2) + 1
		if session("EMPMUN") <> "S" then
          IF selact <> "TODOS" THEN
            strsql = strsql & "WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ") AND NROACTA IS  NULL  "
          ELSE
            strsql = strsql & "WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ")  "
          END IF
        else
        	if selact <> "TODOS" THEN
		       strsql = strsql & "WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ") AND NROACTA IS  NULL and objun=" & session("UN")
		    ELSE
		       strsql = strsql & "WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ") and objun=" & session("UN")
          END IF
        end if

        IF XFNAME = "NO" THEN

            sTitulo = "Esquemas Vigentes Desde: " & Fechador(sID) & " Hasta: " & Fechador(sID2)

        ELSE
            XFSRD = request.querystring("FACTA")
            XFNRO = request.querystring("NACTA")


            sTitulo = " ASIGNADOS AL ACTA NUMERO: " & XFNRO & " DE FECHA " & UCASE(FECHADOR(XFSRD))   & "<BR>"
        END IF


    'strsql = strsql & sverestado


    If Sfiltro <> "TODOS" Then
     strsql = strsql & " and descri='" & Sfiltro & "' "
    End If

     If estado <> "TODOS" Then
		     strsql = strsql & " and estcod=" & estado
    End If

    strsql = strsql & xin

    Select Case sorden
    Case "N"
        strsql = strsql & " order by objdescri, descri"
    Case "T"
        strsql = strsql & " order by  descri,objdescri"
    Case "C"
        strsql = strsql & " order by  fecemis,objdescri, descri"
    Case "V"
        strsql = strsql & " order by  fecvigen,objdescri, descri"
    Case "E"
        strsql = strsql & " order by  estdescri,objdescri, descri"
    End Select


   ' response.write "sql: " & strsql & "<br>"
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")


    Fila = ""

        IF XFNAME = "NO" THEN

           sHTML = OFV.FormHeader(sTitulo)

        ELSE

           sHTML = OFV.FormHeader(session("FormN")) ' ofv.Repbody(UCASE("ESQUEMAS VIGENTES Desde: " & Fechador(sID) & " Hasta: " & Fechador(sID2)))

           XENCABE = ""'ofv.RepHeader2(UCASE("ESQUEMAS VIGENTES  Desde: " & Fechador(sID) & " Hasta: " & Fechador(sID2)),"600")

        END IF



''    sHTML = sHTML & xs

     nHOJA  = 1

     Celda = OFV.GenCelda("", "vt colspan=3", "", "", "", "", "", sTitulo, "si")
     Celda = Celda & OFV.GenCelda("", "VT", "RIGHT", "", "", "", "", "HOJA: " & nHOJA, "si")
     Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


    Celda = ""
    Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "tipo", "si")
    Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "nombre", "si")
    Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "VIGENCIA", "si")
    Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Autor", "si")

    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    Response.write sHTML
    sHTML = ""
    Response.flush

    nLinea = 0
   ' response.write
    fi = CLng(CDate("01-01-1900"))
     PosReg = 0
    'Do Until rs.EOF
  '   On Error Resume Next

  set cn2 = ofv.conectar(ofv.strconn0)

     Do While (Not rs.EOF)
        dEmision = "SD"
        dVigencia = "SD"
        sArchivo = "Iconos/" & rs(11)
        If Not IsNull(rs(0)) Then stipo = rs(0)
        If Not IsNull(rs(5)) Then dEmision = CDate(CLng(rs(5)) + fi)

        dEmi = InStr(dEmision, "1900")
        If dEmi > 0 Then dEmision = "SD"
        If Not IsNull(rs(6)) Then dVigencia = CDate(CLng(rs(6)) + fi)
        dVign = InStr(dVigencia, "1900")
        If dVign > 0 Then dVigencia = "SD"

            sAutor = "&nbsp;"

            If Not IsNull(rs(12)) Then sAutor = rs(12)

        strsql = "select descripcion from usuarios where id = '" & rs(12) & "' "
        Set rs2 = OFV.crearconsultaEx(strsql, cn2, 1, "")
        if not rs2.eof then
           sAutor = rs2(0)
        end if

        If Not IsNull(rs(9)) Then nIden = rs(9)
        If Not IsNull(rs(10)) Then sobjeto = rs(10)

           ' If Not IsNull(rs(19)) Then sobjeto = sobjeto & " - " & rs(19)
            'If Not IsNull(rs(20)) Then sobjeto = sobjeto & " " & rs(20)





        Celda = ""

        Celda = Celda & OFV.GenCelda("", "vt style={font-size:10;}", "", "20%", "", "", "", stipo, "si")
        Celda = Celda & OFV.GenCelda("", "vt style={font-size:10;}", "", "40%", "", "", "", sobjeto, "si")
        Celda = Celda & OFV.GenCelda("", "ut style={font-size:10;}", "", "10%", "", "", "", dVigencia, "si")
        Celda = Celda & OFV.GenCelda("", "ut style={font-size:10;}", "", "30%", "", "", "", sAutor, "si")

        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


        PosReg = PosReg + 1
        rs.movenext
        nLinea = nLinea + 1
      ''  Response.write sHTML
     ''   sHTML = ""
        If nLinea > 30 Then
            nLinea = 0

            IF XFNAME = "ACT" THEN

               XTABLA = OFV.GenTabla("", "", "", "600 rules=rows style={border-style:solid;border-width:1;border-color:#EEEEEE;} ", "", "0", "", "-1", "0", "", "", Fila, "si")

               sHTML = sHTML & "<DIV>" & XENCABE & XTABLA & "</div>"
               Response.write sHTML
               Response.flush
               SHTML = ""

               Fila = ""

               nHOJA  = NHOJA + 1

               Celda = OFV.GenCelda("", "vt colspan=3", "", "", "", "", "", sTitulo, "si")
               Celda = Celda & OFV.GenCelda("", "VT", "RIGHT", "", "", "", "", "HOJA: " & nHOJA, "si")
               Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

               Celda = ""
               Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "tipo", "si")
               Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "nombre", "si")
               Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "VIGENCIA", "si")
               Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "Autor", "si")

               Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


            END IF


        End If

    Loop

    IF XFNAME = "NO" THEN

       sInput = ofv.GenerarInput("","Confirmar","submit","","10","bt","")
       sInput2 = ofv.GenerarInput("ACT1","","READONLY","","20","AA STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:2;BORDER-COLOR:#005500;' ","")
       sInput3 = ofv.GenerarInput("ACT2","","READONLY","","20","AA STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:2;BORDER-COLOR:#005500;' ","")
       Celda = ofv.GenCelda("","vt  colspan=6 ","","","","","","<br><br> TOTAL ESQUEMAS SELECCIONADOS : " & posreg & "&nbsp;&nbsp;&nbsp;" & sInput ,"si")
       Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

       XTABLA = OFV.GenTabla("", "", "", "100% rules=rows style={border-style:solid;border-width:1;border-color:#336699;} ", "", "0", "", "-1", "0", "", "", Fila, "si")

       sAccion = "EstActaConfirmar.asp?Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&selact=" & selact & "&estado=" & estado
       Formulario = ofv.GenForm("ActaLis","",sAccion,"","post","",XTABLA,"si")

       sHTML = sHTML & Formulario

    ELSE

               XTABLA = OFV.GenTabla("", "", "", "600 rules=rows style={border-style:solid;border-width:1;border-color:#EEEEEE;} ", "", "0", "", "-1", "0", "", "", Fila, "si")

               sHTML = sHTML & "<DIV>" & XENCABE & XTABLA & "</div>"
               Response.write sHTML
               SHTML = ""

    END IF


    sHTML = sHTML & "</body></html>"




     session("sperfil")=""


    call ofv.cerrarconn(cn2)
    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

''     Generar = sHTML






end if


response.write sHTML
response.flush

Function Fechador(wkfecha)



      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>