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
    stipoC = Request.querystring("tipoC")

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
    SEXPORT=REQUEST.FORM("EXPORT")
   ' RESPONSE.WRITE REQUEST.FORM("EXPORT")
    SHTMLEXPORT="<TABLE>"


 'RESPONSE.WRITE SESSION("TIPOC") & "<BR>"

    SELECT CASE SESSION("TIPOC")


           CASE  "ACCXESQ"

                 XFDES = REQUEST.FORM("DESDE")  + clng(cdate("01-01-1900"))
                 XFHAS = REQUEST.FORM("HASTA")  + clng(cdate("01-01-1900"))



                 XFSEL = REQUEST.FORM("ESEL")

                 SELECT CASE XFSEL

                        CASE "MAS"

                              XWHERE = "  AND (FINICIO > " & (XFDES - 1) & " AND FINICIO < " & (XFHAS + 1) & ") "
                              XORDER = "ORDER BY CANTF DESC, FULTACC DESC "


                        CASE "MEN"

                              XWHERE = "  AND (FINICIO > " & (XFDES - 1) & " AND FINICIO < " & (XFHAS + 1) & ") "
                              XORDER = "ORDER BY CANTF, FULTACC DESC "

                        CASE "NOL"

                              XWHERE = " AND FINICIO IS NULL "
                              XORDER = " ORDER BY FULTACC DESC "

                 END SELECT

                 strsql = strsql & "SELECT A.CODOBJETO, A.TIPOBJETO, A.FULTACC, A.ACON, AVG(ISNULL(PROM,0)) AS PROMF, SUM(ISNULL(CANT,0)) AS CANTF, "
                 strsql = strsql & " MAX(ESTADO) AS ESTADO , MAX(VERSION) AS VERSION, "
                 strsql = strsql & " MAX(NROVER) AS NROVER, MAX(FECVIGEN) AS FECVIGEN, "
                 strsql = strsql & " MAX(CODIDEN) AS CODIDEN, MAX(DESCRI) AS DESCRI  "
                 strsql = strsql & " FROM OBJALOG A LEFT JOIN (select CODOBJETO, TIPOBJETO, FINICIO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
                 strsql = strsql & " from objalogdet "
                 strsql = strsql & "where tacc = 'CONSEN' "
                 strsql = strsql & "GROUP BY CODOBJETO,TIPOBJETO,FINICIO) B "
                 strsql = strsql & "ON A.CODOBJETO = B.CODOBJETO AND A.TIPOBJETO = B.TIPOBJETO "
                 strsql = strsql & " INNER JOIN OBJETOS C ON A.CODOBJETO = C.CODOBJETO  "
                 strsql = strsql & " INNER JOIN OBJPRTY D ON A.CODOBJETO = D.CODOBJETO  "
                 strsql = strsql & "WHERE A.CODOBJETO > 0  AND ESTADO = 3  " & XWHERE
                 strsql = strsql & "GROUP BY A.CODOBJETO,A.TIPOBJETO,FULTACC, ACON "
                 strsql = strsql & XORDER

               ''  RESPONSE.WRITE STRSQL & "<BR>"



           CASE  "ACCXUSR"




    END SELECT


    CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
    CantRegMover = CDbl(Request.querystring("CantRegMover"))





        sverestado = " and objnivel >= " & Session("SegNivel") & " And Estcod In (9,3)"


        If Request.Form.Count > 0 Then

           sID = Request.Form("Desde")
           sID2 = Request.Form("Hasta")
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


    Addqry  = "&Menu=" & smenu & Addqry
    starget = "estactalis.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry  &"&perfil="&replace(sperfil," ","_")
    sVolver = "estactalis.asp?Menu=" & smenu & sMenuAdd &"&perfil="&replace(sperfil," ","_")
    Addqry  = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")
    Set cn  = OFV.conectar(OFV.strconn2)

    sTitulo = "DOCUMENTOS&nbsp;Vigentes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Fechador(sID) & "&nbsp;&nbsp;al&nbsp;&nbsp;" & Fechador(sID2)

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

    sHTML = sHTML & "<center>"

    Celda = OFV.GenCelda("", "ttt colspan=7", "", "", "", "", "", sTitulo, "si")

    Fila = OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")


    Celda = OFV.GenCelda("", "ttr", "", "3%", "", "", "", "*", "si")
    Celda = Celda & OFV.GenCelda("", "ttr colspan=2", "", "55%", "", "", "", "ESQUEMA NORMATIVO", "si")
    Celda = Celda & OFV.GenCelda("", "ttr", "", "20%", "", "", "", "TIPO", "si")
    Celda = Celda & OFV.GenCelda("", "ttr", "", "10%", "", "", "", "VIGENCIA", "si")
    Celda = Celda & OFV.GenCelda("", "ttr COLSPAN=2 ", "CENTER", "30%", "", "", "", "ACCESOS", "si")

    Fila = fila & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")

   if SEXPORT <> "XLS" THEN
    Response.write sHTML
     sHTML = ""
     Response.flush
   END IF
    nLinea = 0
   ' response.write
    fi = CLng(CDate("01-01-1900"))
     PosReg = 0
    'Do Until rs.EOF
     On Error Resume Next
  ''   If Not rs.EOF Then rs.Move (CantRegMover)
  ''   Do While (Not rs.EOF) And (PosReg < CantRegMostrar)


     Do While (Not rs.EOF)

         sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"

         dVigencia = "SD"
         If Not IsNull(rs(9)) Then dVigencia = fechadorcor( CDate(CLng(rs(9)) + clng(cdate("01-01-1900"))),"")
         dVign = InStr(dVigencia, "1900")
         If dVign > 0 Then dVigencia = "SD"

         sTipobjeto = RS(1)
         sCODobjeto = RS(0)
         sobjeto    = RS(0)

         Strsql  = "Select * From Tipobjeto Where tipobjeto='" & sTipobjeto & "'"
         set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         IF  NOT RSX.EOF THEN sTipobjeto = RSX(1)

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
       '  response.write sactcon

        Celda  = ""
        sLink  = OFV.GenLink("", "UT", sHref, "", "", sFuncion, sImagen, "si")
        Celda  = Celda & OFV.GenCelda("", "UTS COLSPAN=2 ", "", "", "", "", "", sLink, "si")

        IF SEXPORT = "XLS" THEN
         SHTMLEXPORT=SHTMLEXPORT & "<TR><TD>" & POSREG+1 & "</TD>"
         SHTMLEXPORT=SHTMLEXPORT & "<TD>" & sobjeto & "</TD>"
         SHTMLEXPORT=SHTMLEXPORT & "<TD>"  &sTipobjeto & "</TD>"
         SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & dVigencia & "</TD>"
         SHTMLEXPORT=SHTMLEXPORT &  "<TD>"  &RS(5) & "</TD>"
         SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDATE(XFDES),"") & "</TD>"
 		 SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDATE(XFHAS),"") & "</TD></TR>"
 		END IF

        sLink  = OFV.GenLink("", "Ut", sHref, "", "", sFuncion, sobjeto, "si")
        Celda  = Celda & OFV.GenCelda("", "UT", "", "", "", "", "", sLink, "si")

        Celda  = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "",sTipobjeto, "si")


        Celda  = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "", dVigencia, "si")




        Celda  = Celda & OFV.GenCelda("", "UT ", "RIGHT", "", "", "", "", RS(5) & "&nbsp;", "si")
       ' Celda  = Celda & OFV.GenCelda("", "UT ", "RIGHT", "", "", "", "", fechadorcor(XFDES), "si")
       ' Celda  = Celda & OFV.GenCelda("", "UT ", "RIGHT", "", "", "", "", fechadorcor(XFHAS), "si")





        Celda  = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "", sAutor, "si")


        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        PosReg = PosReg + 1
        rs.movenext
        nLinea = nLinea + 1
        if SEXPORT <> "XLS" THEN
        Response.write sHTML
        END IF
        sHTML = ""
        If nLinea > 20 Then
            nLinea = 0
            if SEXPORT <> "XLS" THEN
             Response.flush
            END IF
        End If

    Loop

    if PosReg > 0 then


       Celda = ofv.GenCelda("","ttt  colspan=7 ","","","","",""," TOTAL: " & posreg & " DOCUMENTOS             " & sInput,"si")

    else


       Celda = ofv.GenCelda("","ttt  colspan=7 ","","","","",""," NO HAY DOCUMENTOS PARA LA PRESENTE SELECCION ","si")
    end if

    Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    XTABLA = OFV.GenTabla("", "", "", "95%  ", "", "0", "", "-1", "0", "", "", Fila, "si")


    Formulario = XTABLA

    sHTML = sHTML & Formulario

    sAccion = "COESTDXESQ.asp"
    Celda   = ofv.BotonVolver(sAccion)

    sHTML = sHTML & OFV.botoneraESSX("0", "", "", "", "", "", "", "", "", "", "", "", Celda)

    sHTML = sHTML & "</body></html>"

     session("sperfil")=""


    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

''     Generar = sHTML


if SEXPORT ="XLS" THEN
   Response.ContentType = "application/vnd.ms-excel"

   response.write SHTMLEXPORT & "</TABLE>"
ELSE
   response.write sHTML
   response.flush

END IF


end if




Function Fechador(wkfecha)



      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>