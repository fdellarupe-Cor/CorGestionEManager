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

    NCODIGO2 = REQUEST.querystring("USR")
    NCODIGO = REQUEST.querystring("CODIGO")
    STIPO   = REQUEST.querystring("TIPO")
    SCOD    = REQUEST.querystring("COD")


    XFDES = REQUEST.querystring("FDES")
    XFHAS = REQUEST.querystring("FHAS")

    XWHERE = " AND (FINICIO > " & (XFDES - 1) & " AND FINICIO < " & (XFHAS + 1) & ") AND USUARIO = '" & NCODIGO2 & "' "

    IF SCOD = "SI" THEN
       XWHERE = XWHERE & " AND CODOBJETO = " & NCODIGO & " AND TIPOBJETO = '" & STIPO & "' "
    ELSE
       XWHERE = XWHERE & " AND CODOBJETO > 0 "
    END IF
    XORDER = " ORDER BY USUARIO, CODOBJETO, FINICIO DESC "



                 strsql = strsql & "SELECT CODOBJETO, TIPOBJETO, FINICIO, HINICIO, HFINAL , "
                 strsql = strsql & "(HFINAL-HINICIO) AS TIEMPO,"
                 strsql = strsql & " USUARIO, USUARIO AS USERACC "
                 strsql = strsql & " FROM OBJALOGDET B1"
                 strsql = strsql & " WHERE TACC = 'CONSEN'  " & XWHERE

                 strsql = strsql & XORDER

                '' RESPONSE.WRITE STRSQL & "<BR>"





    sID  = XFDES  - clng(cdate("01-01-1900"))
    sID2 = XFHAS  - clng(cdate("01-01-1900"))

    sTitulo = ""


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


 '   Celda = OFV.GenCelda("", "utg colspan=2 style={border-style:RIDGE;border-width:2;border-color:#F7AA22;} ", "", "", "", "", "", sTitulo, "si")

'    Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



     xxxx    = getInfoUser(NCODIGO2)
     sGrpUsr = session("SGU_nombre")


    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "USUARIO:", "si")
    Celda = Celda & OFV.GenCelda("", "uti", "", "80%", "", "", "", "" & sGrpUsr & "", "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



 '   Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "SECTOR:", "si")
 '   Celda = Celda & OFV.GenCelda("", "UTi", "", "80%", "", "", "", RSX(1), "si")
 '   Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "DESDE:", "si")
    Celda = Celda & OFV.GenCelda("", "UTi", "", "80%", "", "", "", UCASE(Fechador(sID)), "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    Celda = OFV.GenCelda("", "VT ", "","20%", "", "", "", "HASTA:", "si")
    Celda = Celda & OFV.GenCelda("", "UTi", "", "80%", "", "", "", UCASE(Fechador(sID2)), "si")
    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

   '' sTitulo = "ACCESO DE USUARIOS<BR> Desde: " & Fechador(sID) & "<BR> Hasta: " & Fechador(sID2)





    XTABLA = OFV.GenTabla("", "aaa", "", "95%", "", "0", "", "0", "0", "", "", Fila, "si")

    Fila = ""

    Celda = OFV.GenCelda("", "vt colspan=5", "", "", "", "", "", " ", "si")

    Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


    Celda = ""
    Celda = Celda & OFV.GenCelda("", "TTr", "","50%", "", "", "", "DOCUMENTO", "si")
    Celda = Celda & OFV.GenCelda("", "TTr", "","25%", "", "", "", "FECHA", "si")
    Celda = Celda & OFV.GenCelda("", "TTr", "","25%", "", "", "", "ACCESOS", "si")

    Fila = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    XTABLA2 = OFV.GenTabla("", "aaa", "", "100%", "", "0", "", "0", "0", "", "", Fila, "si")

    sHTML = sHTML & "<CENTER>"
    sHTML = sHTML & "<div id='mt1' class='DINDM' style='display:block;width:650;' >"
    sHTML = sHTML & XTABLA & XTABLA2
    sHTML = sHTML & "</div>"

    Call OFV.cerrarconsulta(rsX)

    sHTML = sHTML & "<div id='mt2' class='DINDM' style='display:block;height:300px;width:650;overflow:auto;' >"


    Fila = ""




    Celda = OFV.GenCelda("", "UT colspan=5", "", "", "", "", "", " ", "si")

    Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



    nLinea = 0


     PosReg = 0
     Postmp = 0
     Poscnt = 0

     PosReg2 = 0
     Postmp2 = 0


     SUSRANT = "NO"



   ' response.write "sql: " & strsql & "<br>"
    Set rsX = OFV.crearconsultaEx(strsql, cn, 1, "")
     Do UNTIL rsX.EOF
    '    RESPONSE.WRITE "U: " & RSX(6) & RSX(2) & "<BR>"

           sImagen  = "<font face=wingdings size=4 color=#993333>2</font>"
           sImagen2 = "<font face=wEBdings  size=2 color=#F7AA22>4</font>"

         IF SUSRANT <> cstr(RSX(0) & "") THEN

            xTrColor = "#F9F9F9"

            IF SUSRANT <> "NO" THEN

               Celda = OFV.GenCelda("", "Ut COLSPAN=2 ", "RIGHT", "", "", "", "","&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "TOTAL&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "ACCESOS:&nbsp;" & PosReg2 & "&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "PROMEDIO:&nbsp;" & GETGLBHMS(round((Postmp2/PosReg2),0)) & "&nbsp;", "si")

               Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

            end if

            PosReg2 = 0
            Postmp2 = 0
            Poscnt  = Poscnt + 1

            SUSRANT = cstr(RSX(0))

         sTipobjeto = RSX(0)

         Strsql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & sTipobjeto
         set rsM = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         IF NOT RSM.EOF THEN
           sTipobjeto1 = RSM(1) & ":&nbsp;"
           sTipobjeto2 = RSM(3) & " - "
           sTipobjeto3 = RSM(2) & " ("
           sTipobjeto4 = RSM(5) & " "
           sTipobjeto5 = RSM(6) & ") "
           sTipobjeto  = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
         END IF

            Celda = OFV.GenCelda("", "UTS", "", "2%", "", "", "", sImagen, "si")
            Celda = Celda & OFV.GenCelda("", "VT COLSPAN=4 ", "", "", "", "", "", "&nbsp;" & sTipobjeto, "si")

            Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")
         END IF


         dVigencia = "SD"
         If Not IsNull(rsX(2)) Then dVigencia = CDate(CLng(rsX(2)))
         dVign = InStr(dVigencia, "1900")
         If dVign > 0 Then dVigencia = "SD"


         XTIME =  RSX(5) & ""

         IF XTIME = "" OR XTIME = "0" THEN XTIME = "1"

         XTIME =  CLNG(XTIME)


        Celda = OFV.GenCelda("", "AAA COLSPAN=2 ", "RIGHT", "", "", "", "","&nbsp;", "si")
        Celda = Celda & OFV.GenCelda("", "UTS", "RIGHT", "", "", "", "",sImagen2, "si")
        Celda = Celda & OFV.GenCelda("", "UT  bgcolor='" & xTrColor & "' ", "RIGHT", "", "", "", "", FechadorCor(dVigencia,0) & "&nbsp;", "si")
        Celda = Celda & OFV.GenCelda("", "UT  bgcolor='" & xTrColor & "' ", "RIGHT", "", "", "", "", GETGLBHMS(XTIME) & "&nbsp;", "si")

        Fila = Fila & OFV.GenRow("", "AAA", "", "", "", "", "", Celda, "si")

        if  xTrColor = "#F9F9F9" then
            xTrColor = "#F3F3F3"
        ELSE
            xTrColor = "#F9F9F9"
        END IF



        PosReg = PosReg + 1
   '     PosReg2 = PosReg2 + RSX(8)
        PosReg2 = PosReg2 + 1

        Postmp = Postmp + XTIME

        Postmp2 = Postmp2 + XTIME

        rsX.movenext


    Loop




    if PosReg > 0 then



            IF SUSRANT <> "NO" THEN

               Celda = OFV.GenCelda("", "vt COLSPAN=2 ", "RIGHT", "", "", "", "","&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "TOTAL&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "ACCESOS:&nbsp;" & PosReg2 & "&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "PROMEDIO:&nbsp;" & GETGLBHMS(round((Postmp2/PosReg2),0)) & "&nbsp;", "si")

               Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

            end if



               Celda = OFV.GenCelda("", "vt COLSPAN=2 ", "", "", "", "", "","TOTAL GENERAL&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "DOCUMENTOS:&nbsp;" & Poscnt, "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "ACCESOS:&nbsp;" & PosReg & "&nbsp;", "si")
               Celda = Celda & OFV.GenCelda("", "UT", "right", "", "", "", "", "PROMEDIO:&nbsp;" & GETGLBHMS(round((Postmp/PosReg),0)) & "&nbsp;", "si")

               FilaF = OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")
               XTABLAF = OFV.GenTabla("", "aaa", "", "95%", "", "0", "", "-1", "0", "", "", FilaF, "si")



    else

       Celda = ofv.GenCelda("","vt  colspan=5 ","","","","",""," NO HAY ACCESOS PARA LA PRESENTE SELECCION ","si")
       Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    end if

    Celda = OFV.GenCelda("", "UT", "", "2%", "", "", "","&nbsp;", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "43%", "", "", "", "&nbsp;", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "10%", "", "", "", "&nbsp;", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "20%", "", "", "", "&nbsp;", "si")
    Celda = Celda & OFV.GenCelda("", "UT", "", "25%", "", "", "", "&nbsp;", "si")
    Fila  = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


    XTABLA = OFV.GenTabla("", "aaa", "", "95%", "", "0", "", "-1", "0", "", "", Fila, "si")
    sHTML  = sHTML & XTABLA
    sHTML  = sHTML &  "</div>"

    sHTML = sHTML & "<div id='mtF' class='DINDM' style='display:block;width:650;' >"
    sHTML  = sHTML & XTABLAf
    sHTML  = sHTML &  "</div>"
    sHTML  = sHTML & "</CENTER>"
    sHTML  = sHTML &  "</body></html>"

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