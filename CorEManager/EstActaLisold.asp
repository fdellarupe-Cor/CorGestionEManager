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



    CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
    CantRegMover = CDbl(Request.querystring("CantRegMover"))


      'response.write Request.Form("selact") & chr(13) & request.querystring


        sverestado = " and objnivel >= " & Session("SegNivel") & " And Estcod In (9,3)"


 '   For i = 1 To 4
 '       If bconsultar(i) Then sverestado = sverestado & i & ","
 '   Next
 '   sverestado = Left(sverestado, Len(sverestado) - 1) & ") "



        If Request.Form.Count > 0 Then

           sID = Request.Form("Desde")
           sID2 = Request.Form("Hasta")
           sID3 = Request.Form("Fecha")
           selact=UCASE(Request.Form("selact") &"")
           Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&selAct=" & selact
        Else
           sID = Request.querystring("Desde")
           sID2 = Request.querystring("Hasta")
           sID3 = Request.querystring("Fecha")
           selact=Request.querystring("selact")
           Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 &  "&selAct=" & selact
        End If


       '' response.write "fd " & sid & " fh " & sid2 & " ts " & sid3 & " <br>"



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
            if selact <> "TODOS" THEN
             strsql = strsql & "WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ") AND NROACTA IS  NULL "
            ELSE
             strsql = strsql & "WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ")  "
            END IF
         ''   strsql = strsql & "WHERE xfecvigen between '" & cdesde & "' and '" & chasta & "' "
            sTitulo = "Esquemas Vigentes Desde: " & Fechador(sID) & " Hasta: " & Fechador(sID2)


    strsql = strsql & sverestado


    If Sfiltro <> "TODOS" Then
     strsql = strsql & " and descri='" & Sfiltro & "' "
    End If

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

    sHTML = OFV.FormHeader(sTitulo)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs


    Celda = OFV.GenCelda("", "vt colspan=6", "", "", "", "", "", sTitulo, "si")

    Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


    Celda = OFV.GenCelda("", "TTR", "", "", "", "", "", "*", "si")
    Celda = Celda & OFV.GenCelda("", "TTR colspan=2", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
    Celda = Celda & OFV.GenCelda("", "TTR", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=T target=_self ><B><u>&nbsp;TIPO&nbsp;</u></B></A>", "si")
    Celda = Celda & OFV.GenCelda("", "TTR", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=V target=_self ><B><u>&nbsp;VIGENCIA&nbsp;</u></B></A>", "si")
   ' Celda = Celda & OFV.GenCelda("", "TTR", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CREACION&nbsp;</u></B></A>", "si")
    Celda = Celda & OFV.GenCelda("", "TTR", "", "", "", "", "", "Autor", "si")

    Fila = fila & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")

    Response.write sHTML
    sHTML = ""
    Response.flush

    nLinea = 0
   ' response.write
    fi = CLng(CDate("01-01-1900"))
     PosReg = 0
    'Do Until rs.EOF
     On Error Resume Next
  ''   If Not rs.EOF Then rs.Move (CantRegMover)
  ''   Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

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

        If Not IsNull(rs(9)) Then nIden = rs(9)
        If Not IsNull(rs(10)) Then sobjeto = rs(10)
   ''     If smenu <> "PORFECHA" Then
            If Not IsNull(rs(19)) Then sobjeto = sobjeto & " - " & rs(19)
            If Not IsNull(rs(20)) Then sobjeto = sobjeto & " " & rs(20)
  ''      End If
        If Session("USRES") <> "B" Then
           sImagen = OFV.GenImage("", "", "", "", sArchivo, "0", "")
        Else
           sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"
        End If

        sActcon = "Consultar.asp?Codigo=" & rs("codobjeto") & "&Tipo=" & rs("tipobjeto")
        sActcon = sActcon & "&Perfil=" & rs("perfil") & "&Datos=" & OFV.convertircar(sMenuMas, "&", "*")
        sActcon = sActcon & "&Volver=" & sVolver
        sHref = "&nbsp;"
        sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)
       '  response.write sactcon

        SIMGSEL = "<A id='I" & rs(3) & "' HREF='Javascript:void(0);' ONCLICK=" & chr(34) & "this.style.display='none';document.all.E" & rs(3) & ".style.display='block';document.all.M" & rs(3) & ".value='0';" & chr(34) & " STYLE='DISPLAY:BLOCK;' title='Excluir' ><FONT FACE=corflaticon SIZE=3 COLOR=#008000>5</FONT></A> "
        SIMGSEL = SIMGSEL & "<A id='E" & rs(3) & "' HREF='Javascript:void(0);' ONCLICK=" & chr(34) & "this.style.display='none';document.all.I" & rs(3) & ".style.display='block';document.all.M" & rs(3) & ".value='1';" & chr(34) & " STYLE='DISPLAY:none;' title='Incluir' ><FONT FACE=corflaticon SIZE=3 COLOR=#aa0000>6</FONT></A> "
        SIMGSEL = SIMGSEL & ofv.GenerarInput("M" & rs(3),"1","hidden","","1","dt","")

       ' SIMGSEL = "<A id='I" & rs(3) & "' HREF='Javascript:void(0);' ONCLICK=" & chr(34) & "this.style.display='none';document.all.E" & rs(3) & ".style.display='block';document.all.M" & rs(3) & ".value='1';" & chr(34) & " STYLE='DISPLAY:BLOCK;' title='Incluir' ><FONT FACE=corflaticon SIZE=3 COLOR=#aa0000>6</FONT></A> "
	  '  SIMGSEL = SIMGSEL & "<A id='E" & rs(3) & "' HREF='Javascript:void(0);' ONCLICK=" & chr(34) & "this.style.display='none';document.all.I" & rs(3) & ".style.display='block';document.all.M" & rs(3) & ".value='0';" & chr(34) & " STYLE='DISPLAY:none;' title='Excluir' ><FONT FACE=corflaticon SIZE=3 COLOR=#008000>5</FONT></A>  "
      '  SIMGSEL = SIMGSEL & ofv.GenerarInput("M" & rs(3),"0","hidden","","0","dt","")


        Celda = OFV.GenCelda("", "", "", "", "", "", "", sIMGSEL, "si")

        sLink = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen, "si")

        Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", sLink, "si")

        sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, sobjeto, "si")

        Celda = Celda & OFV.GenCelda("", "abc style={font-size:10;}", "", "", "", "", "", sLink, "si")
        Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=" & sorden & "&sfiltro=" & Replace(stipo, " ", "_") & " target=_self ><font face=verdana size=1 color=#336699><b>" & stipo & "</b></font></A>", "si")

        Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", dVigencia, "si")
     '   Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", dEmision, "si")

        Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", sAutor, "si")


        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        PosReg = PosReg + 1
        rs.movenext
        nLinea = nLinea + 1
        Response.write sHTML
        sHTML = ""
        If nLinea > 20 Then
            nLinea = 0
            Response.flush
        End If

    Loop

    if PosReg > 0 then

       sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
       Celda = ofv.GenCelda("","vt  colspan=6 ","","","","",""," TOTAL: " & posreg & " ESQUEMAS             " & sInput,"si")

    else


       Celda = ofv.GenCelda("","UT  colspan=6 ","","","","","","&nbsp;&nbsp;<b><font face=""corflaticon"" size=4 color=""#993333"" >&#0105;</font></b>&nbsp;&nbsp;<font size=2 >NO HAY DOCUMENTOS PARA LA PRESENTE SELECCION</font>","si")
    end if

    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    XTABLA = OFV.GenTabla("", "", "", "100% rules=rows style={border-style:solid;border-width:1;border-color:#336699;} ", "", "0", "", "-1", "0", "", "", Fila, "si")

    sAccion = "EstActaLisConf.asp?Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0"
    Formulario = ofv.GenForm("ActaLis","",sAccion,"","post","",XTABLA,"si")

    sHTML = sHTML & Formulario & "</body></html>"


   ''  sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry  & "&perfil="&replace(sperfil," ","_"), "", "no", "", "", "")

     session("sperfil")=""


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