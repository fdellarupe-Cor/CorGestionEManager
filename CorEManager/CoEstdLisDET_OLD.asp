<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else



    Dim bconsultar()

    

    strsql = ""
    strsql1 = ""

    NCODIGO = REQUEST.querystring("CODIGO")
    STIPO   = REQUEST.querystring("TIPO")

 '   XFDES = REQUEST.querystring("FDES")  - clng(cdate("01-01-1900"))
 '   XFHAS = REQUEST.querystring("FHAS")  - clng(cdate("01-01-1900"))

    XFDES = REQUEST.querystring("FDES")
    XFHAS = REQUEST.querystring("FHAS")

    XWHERE = " AND (FINICIO > " & (XFDES - 1) & " AND FINICIO < " & (XFHAS + 1) & ") AND CODOBJETO = " & NCODIGO
    XWHERE = XWHERE & "  AND TIPOBJETO = '" & STIPO & "' "
    XORDER = " ORDER BY  USUARIO, FINICIO DESC " 




                 strsql = strsql & "SELECT CODOBJETO, TIPOBJETO, FINICIO, MAX(HINICIO), MAX(HFINAL) , "
                 strsql = strsql & "AVG(HFINAL-HINICIO) AS PROMF,"
                 strsql = strsql & " USUARIO, USUARIO AS USERACC , SUM(1) AS CANT"
                 strsql = strsql & " FROM OBJALOGDET B1"
                 strsql = strsql & " WHERE TACC = 'CONSEN'  " & XWHERE
                 strsql = strsql & " GROUP BY CODOBJETO, TIPOBJETO, USUARIO, FINICIO"
                 strsql = strsql & XORDER

                '' RESPONSE.WRITE STRSQL & "<BR>"

                 strsql1 = strsql1 & "SELECT DESCRI, CODIDEN, FECVIGEN, ESTADO, VERSION , "
                 strsql1 = strsql1 & " ISNULL(NROVER,'0') AS NROVER,FULTACC, B.AUTOR AS USERAUT "
                 strsql1 = strsql1 & " FROM OBJETOS A "
                 strsql1 = strsql1 & " INNER JOIN OBJPRTY B ON A.CODOBJETO = B.CODOBJETO "
                 strsql1 = strsql1 & " LEFT JOIN OBJALOG C ON A.CODOBJETO = C.CODOBJETO "
                ''''' strsql1 = strsql1 & " LEFT JOIN USUARIOS D ON B.AUTOR = D.ID "
                 strsql1 = strsql1 & " WHERE A.CODOBJETO = " & NCODIGO


                 ''RESPONSE.WRITE STRSQL1 & "<BR>"



    sID  = XFDES  - clng(cdate("01-01-1900"))
    sID2 = XFHAS  - clng(cdate("01-01-1900"))

  '  sID  = XFDES
   ' sID2 = XFHAS

    sTitulo = "ACCESO DE USUARIOS<BR> Desde: " & Fechador(sID) & "<BR> Hasta: " & Fechador(sID2)


    Set cn = OFV.conectar(OFV.strconn2)



    fi = CLng(CDate("01-01-1900"))



    sHTML = OFV.FormHeader(sTitulo)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs


    Celda = OFV.GenCelda("", "vt colspan=2 style={border-style:RIDGE;border-width:2;border-color:#F7AA22;} ", "", "", "", "", "", sTitulo, "si")

    Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    Set rsX = OFV.crearconsultaEx(strsql1, cn, 1, "")

    IF NOT RSX.EOF THEN


    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "<B><u>&nbsp;DOCUMENTO:&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "VT", "", "80%", "", "", "", "<BIG>" & RSX(0) & " - " & RSX(1) & "</BIG>", "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

         sTipobjeto = RSX(3)

         Strsql2 = "Select * From ESTADOS Where OBJID = 0" & sTipobjeto
         set rsM = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         IF NOT RSM.EOF THEN sTipobjeto = RSM(1)

        xxxx    = getInfoUser(RSX(7))
        sGrpUsr = session("SGU_nombre")



    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "<B><u>&nbsp;AUTOR:&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "80%", "", "", "", sGrpUsr, "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "<B><u>&nbsp;ESTADO:&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "80%", "", "", "", sTipobjeto, "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

         dVigencia = "SD"
         If Not IsNull(rsX(2)) Then dVigencia = CDate(CLng(rsX(2)) + clng(cdate("01-01-1900")))
         dVign = InStr(dVigencia, "1900")
         If dVign > 0 Then dVigencia = "SD"

    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "<B><u>&nbsp;VIGENCIA:&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "80%", "", "", "", dVigencia, "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

         dVigencia = "SD"
         If Not IsNull(rsX(6)) Then dVigencia = CDate(CLng(rsX(6)))
         dVign = InStr(dVigencia, "1900")
         If dVign > 0 Then dVigencia = "SD"

    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "<B><u>&nbsp;ULTIMO ACCESO:&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "80%", "", "", "", dVigencia, "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




    END IF


    XTABLA = OFV.GenTabla("", "", "", "80% style={border-style:solid;border-width:1;border-color:#336699;} ", "", "0", "", "0", "0", "", "", Fila, "si")

    Formulario = XTABLA

    sHTML = sHTML & "<BR><BR><CENTER>" &  Formulario & "</CENTER><BR>"
    
    Call OFV.cerrarconsulta(rsX)


    sHTML = sHTML & "<center><div id='mt2' class='wwwDINDM' style='display:block;height:260px;width:650;overflow:auto;' ></center>"

    Celda = OFV.GenCelda("", "vt colspan=5", "", "", "", "", "", "DETALLE DE ACCESOS", "si")

    Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


    Celda = OFV.GenCelda("", "tt", "", "3", "", "", "", "*", "si")
    Celda = Celda & OFV.GenCelda("", "TT ", "","70%", "", "", "", "<B><u>&nbsp;USUARIO&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "TT", "", "10%", "", "", "", "<B><u>&nbsp;FECHA&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "TT", "", "10%", "", "", "", "<B><u>&nbsp;ACCESOS&nbsp;</u></B>", "si")
    Celda = Celda & OFV.GenCelda("", "TT", "", "10%", "", "", "", "<B><u>&nbsp;PROMEDIO&nbsp;</u></B>", "si")

    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    nLinea = 0


     PosReg = 0
     PosReg2 = 0
     SUSRANT = "NO"

   '' response.write "sql: " & strsql & "<br>"
    Set rsX = OFV.crearconsultaEx(strsql, cn, 1, "")
     Do UNTIL rsX.EOF 
    '    RESPONSE.WRITE "U: " & RSX(6) & RSX(2) & "<BR>"    

           sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"

         IF SUSRANT <> (RSX(6) & "") THEN
            SUSRANT = RSX(6)

            xxxx    = getInfoUser(RSX(6))
           sGrpUsr = session("SGU_nombre")

            Celda = OFV.GenCelda("", "", "", "", "", "", "", "<FONT FACE=corflaticon SIZE=3 COLOR=#008000>&#0128;</FONT>", "si")
            Celda = Celda & OFV.GenCelda("", "VT COLSPAN=4 ", "", "", "", "", "", "&nbsp;" & sGrpUsr, "si")

            Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
         END IF


         dVigencia = "SD"
         If Not IsNull(rsX(2)) Then dVigencia = CDate(CLng(rsX(2)))
         dVign = InStr(dVigencia, "1900")
         If dVign > 0 Then dVigencia = "SD"


        Celda = OFV.GenCelda("", "AAA COLSPAN=2 ", "RIGHT", "", "", "", "","&nbsp;", "si")
        Celda = Celda & OFV.GenCelda("", "UT", "", "", "", "", "", dVigencia, "si")
        Celda = Celda & OFV.GenCelda("", "UT", "RIGHT", "", "", "", "", RSX(8) & "&nbsp;", "si")

        
        Celda = Celda & OFV.GenCelda("", "UT", "RIGHT", "", "", "", "", GETGLBHMS(RSX(5)) & "&nbsp;", "si")

        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        PosReg = PosReg + 1
        PosReg2 = PosReg2 + RSX(8)
        rsX.movenext

        
    Loop

    if PosReg > 0 then


       Celda = ofv.GenCelda("","vt style={border-style:RIDGE;border-width:1;border-color:#F7AA22;} colspan=5 ","","","","",""," TOTAL: " & posreg2 & " ACCESOS             " & sInput,"si")

    else

       Celda = ofv.GenCelda("","vt  colspan=5 ","","","","",""," NO HAY ACCESOS PARA LA PRESENTE SELECCION ","si")

    end if

    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    XTABLA = OFV.GenTabla("", "", "", "100% style={border-style:solid;border-width:1;border-color:#336699;} ", "", "0", "", "-1", "0", "", "", Fila, "si")

    Formulario = XTABLA

    sHTML = sHTML & "<CENTER>" &  Formulario & "</CENTER></div></body></html>"
    
    Call OFV.cerrarconsulta(rsX)
    Call OFV.cerrarconn(cn)



end if


response.write sHTML


Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>