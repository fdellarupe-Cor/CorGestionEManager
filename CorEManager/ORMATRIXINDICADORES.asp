<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = ofv.CheckUsuario()
IF ok <> "Y" THEN
   uv = ofv.uv
   IF uv = "NO" THEN ok = ofv.ValidarUsuario
   sHTML = ok
ELSE
   response.expires = 0


   CantRegMostrarx = 8
   CantRegMoverx = CDBL(request.querystring("CantRegMover"))
   vuelta = ofv.vueltaasp("../", CantRegMostrarx, Cantregmoverx)


   SORCODIGO = request.FORM("ORCODIGO")
   sORTEXTO = request.FORM("ORTEXTO")
   SORESTADO = request.FORM("ORESTADO")
   sORCLASIFICAINDI = request.FORM("ORCLASIFICAINDI")
   SORVENCIDOS = request.FORM("ORVENCIDOS")
   SORUMBRAL=request.FORM("ORUMBRAL")
   SFRECUENCIA=request.FORM("ORFRECUENCIA")
  ' RESPONSE.WRITE SORUMBRAL

   WSPSESION = "AND"
   SESSION(WSPSESION) = ""

   IF SORCODIGO <> "TODOS" THEN

      SORCODIGOW = " (A.OBJID = 0" & SORCODIGO & ") "

      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORCODIGOW
      SCONECTOR = " AND "

   END IF

   IF SORTEXTO <> "TODOS" THEN
      sORCARACTERW = " (A.DESCRIPCION LIKE '%" & SORTEXTO & "%') "

      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & sORCARACTERW
      SCONECTOR = " AND "

   END IF


   IF SORESTADO = "VENC" THEN
      SORTEXTOW = " (A.ESTADO=2) "

      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
      SCONECTOR = " AND "

   ELSE
      IF SORESTADO <> "TODOS" THEN
         SORTEXTOW = " (A.ESTADO="& SORESTADO &") "

         SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
         SCONECTOR = " AND "

      END IF
   END IF


   IF sORCLASIFICAINDI <> "TODOS" THEN
      SORTEXTOW = " (A.CLASIFICACION="& sORCLASIFICAINDI &") "
      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
      SCONECTOR = " AND "

   END IF


   IF SFRECUENCIA <> "TODOS" THEN
         SORTEXTOW = " (A.UNIDADFRECUEN= "& SFRECUENCIA &") "
         SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
         SCONECTOR = " AND "

   END IF







   sHTML = ofv.FormHeader(session("FormN"))


   ofv.ObtenerAtributos Session("Form"), ""
   bConsultar = ofv.AConsultar

   SET cn = ofv.conectar(ofv.strconn2)



   Fila = ""
   xTrColor = "#FFFFFF"

   sImageva = "<font face='corflaticon' size=2 color='#990000'>&#0108;</font>"
   sImagevaC = "&#0108;"
   sImagevm = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
   sImagevb = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
   sImagevn = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

   sImagenO  = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
   sImagenI  = "<B><FONT FACE=corflaticon COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
   sImagenE  = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
   sImagenC  = "<B><FONT FACE=corflaticon COLOR=#008000 SIZE=3>R</FONT></B>"
   sImagenP  = "<B><FONT FACE=corflaticon COLOR=#336699 SIZE=3>2</FONT></B>"
   sImagenP1  = "<B><FONT FACE=corflaticon COLOR=#336699 SIZE=3>2</FONT></B>"
   sImagenRU = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
   sImagenRG = "<B><FONT FACE=corflaticon COLOR=#00aa00 SIZE=3>&#0148;</FONT></B>"
   sImagenRS = "<B><FONT FACE=corflaticon COLOR=#0000aa SIZE=3>G</FONT></B>"
   sImagenFI = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>5</FONT></B>"
   sImagenFF = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>6</FONT></B>"

   sImagenA  = "<FONT FACE=corflaticon COLOR=#005500 SIZE=3>&#0252;</FONT>"
   sImagenR  = "<FONT FACE=corflaticon COLOR=#AA0000 SIZE=4>&#0251;</FONT>"
   sImagenK  = "<FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0184;</FONT>"
   sImagenIN = "<FONT FACE=corflaticon COLOR=#336699 SIZE=3>&#0236;</FONT>"
   flechaup= "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=5>&#0241;</FONT></B>"
   flechadown= "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=5>&#0242;</FONT></B>"
   flechaequal= "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=5>&#0243;</FONT></B>"
   sIFORM = "<font face='corflaticon' size=2 color='#FFFFFF' onclick=""document.all.DX5.style.display='block';"" TITLE='Ver Formula de aplicacion.' style='cursor:pointer;background-color:#000099;border-style:OUTSET;border-width:1;' >i</font>"

   wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "

   ''''b.codvinobj = 'VN' AND

   BCOLORG = "#AAAAAA"
   BCOLORD = "#6699CC"

   MAINPROC = ""
   MAINACTI = ""
   MAINRIES = ""

   WKTITG = "MATRIZ DE INDICADORES  DE RIESGOS" & sImagenP

   wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "


   Celda = ""
   Celda = Celda & ofv.gencelda("", "VTI COLSPAN=8 " & wkstylo1, "", "", "", "", "", WKTITG, "si")


   Fila = FILA & ofv.genrow("", "AAA BGCOLOR='#F7EEDD' ", "", "", "25", "", "", Celda, "si")



   Celda = ""

   tabla1 = ofv.GenTabla("", "aaa  STYLE='border-style:solid;border-width:3;border-color:" & BCOLORG & ";' ", "", "95% ", "", "0", "", "1", "1", "", "", Fila, "si")
   FILA = ""


   CELDAINFO =  ofv.gencelda("", "VTI ALIGN=CENTER " , "", "", "", "", "", "NOMBRE INDICADOR", "si")
   CELDAINFO = CELDAINFO & ofv.gencelda("", "VTI ALIGN=CENTER " , "", "", "", "", "", "VENCIDO", "si")
   CELDAINFO = CELDAINFO & ofv.gencelda("", "VTI ALIGN=CENTER " , "", "", "", "", "", "UMBRAL", "si")
   CELDAINFO = CELDAINFO & ofv.gencelda("", "VTI ALIGN=CENTER " , "", "", "", "", "", "VALOR", "si")
   CELDAINFO = CELDAINFO & ofv.gencelda("", "VTI ALIGN=CENTER " , "", "", "", "", "", "RIESGOS ASOC.", "si")
   CELDAINFO = CELDAINFO & ofv.gencelda("", "VTI ALIGN=CENTER " , "", "", "", "", "", "USR RESP", "si")
   CELDAINFO = CELDAINFO & ofv.gencelda("", "VTI ALIGN=LEFT" , "", "", "", "", "", "ESTADO", "si")
   CELDAINFO = CELDAINFO & ofv.gencelda("", "VTI ALIGN=CENTER " , "", "", "", "", "", "TENDENCIA", "si")
   FILAinfo =  ofv.genrow("AAA BGCOLOR='#F7EEDD'", "", "", "", "", "", "", CELDAINFO, "si")


   IF session(WSPSESION )<> "" THEN
      swhere = " where "
 '  ELSE
'      swhere = "where  a.estado=2"
   END IF

   'STRSQL = "select * , (select count(*) from ORRELAINDIRIESGO b where b.IDINDICADOR=a.objid ) as REL ,(select c.valor from orfrecuencia c where c.objid=a.unidadfrecuen) as dias FROM ORINDICADORES a  " & swhere & SESSION(WSPSESION)
   STRSQL = "select * , (select count(*) from ORRELAINDIRIESGO b where b.IDINDICADOR=a.objid ) as REL ,(select c.valor from orfrecuencia c where c.objid=a.unidadfrecuen) as dias FROM ORINDICADORES a  " & swhere & SESSION(WSPSESION)
    'response.write strsql & SORESTADO

   SET rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   IF NOT rs.EOF THEN

      posi = 0
      coln = 0
      Celda = ""
      fila = ""
      xTrColor = "#EEEEEE"
      DO WHILE NOT RS.EOF

         SVAUMBRAL="NO"
         strsql = "select *, " 'from csorvariables3 where varnombre='Z' and idindicador=0" & rs(0)
         strsql = strsql & "(SELECT B.VALOR  FROM ORVARIABLES B WHERE B.VARNOMBRE='Z'  AND B.ACTIVA=1 AND  B.IDINDICADOR=0"& rs(0) &" AND  B.grupo=isnull(a.GRUPO,0 ) -1 ) AS  ultresult"
         strsql =strsql & " from csorvariables3 a where varnombre='Z' and idindicador=0" & rs(0)
        ' response.write strsql

         SET vars = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

         sinputVal = validarFecuencia2(rs(0), rs("FECHAULTACT"), rs("dias"))


         SELECT CASE SORESTADO

            CASE "VENC"
               IF sinputVal = "SI" THEN

                  IF NOT vars.EOF THEN
                     VALOR = REPLACE(vars("VALOR")&"",",",".")
                     FECHACT = vars("FECHACT")
                     ultresul=vars("ultresult")
                  ELSE
                     VALOR = 0
                     FECHACT = 0
                     ultresul=0
                  END IF


                  IF CDBL("0" & VALOR) = 0 AND CLNG("0" & FECHACT) = 0 THEN
                     descripcionUM = "NO EXISTEN VARIABLES"
                     sImageva = "<font face='corflaticon' size=6 color='#ff0000'>"&sImagevaC&"</font>"
                  ELSE

                     Strsql = "Select * From orumbrales  WHERE VALOR1 <=0"&VALOR& " AND VALOR2>=0"& VALOR & " AND IDINDICADOR="&rs(0) 'id
                     '
                     '  response.write strsql & "<br>"
                      SVAUMBRALFUERA="NO"
                     SET rsum = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

                     IF rsum.EOF THEN
                        '
                        sImageva = "<font face='corflaticon' size=6 color='#ff0000'>"&sImagevaC&"</font>"
                        descripcionUM = "FUERA DE RANGOS"
                        SVAUMBRALFUERA="SI"
                     ELSE

                        sImageva = "<font face='corflaticon' size=6 color='" & rsum("color") & "'>"&sImagevaC&"</font>"
                        descripcionUM = rsum("descripcion")
                        IF CSTR(RSUM("ORDEN")&"")= CSTR(SORUMBRAL&"")  THEN  SVAUMBRAL="SI"
                     END IF
                  END IF


                  CELDAINFOSEM = ofv.gencelda("", "ut   align=CENTER" , "", "", "", "", "", sImageva, "si")

                  info = "<u><b>CODIGO:" & RS(0)&"&nbsp;&nbsp;"& rs("descripcion") & "</b></u>"
                  CELDAINFO =  ofv.gencelda("", "ut " , "", "", "", "", "", info, "si")
                  'info = "<b>" GTFechadorSH(FECHACT)& "</b>"

                  info = "<b>" & sinputVal & "</b>"
                  CELDAINFO =CELDAINFO &  ofv.gencelda("", "ut  align=CENTER" , "", "", "", "", "", info, "si")









				 info =  descripcionUM
				 CELDAINFO = CELDAINFO & ofv.gencelda("", "ut  ALIGN=CENTER" , "", "", "", "", "", info, "si")
				 info =  VALOR
				 CELDAINFO = CELDAINFO & ofv.gencelda("", "ut  ALIGN=RIGHT" , "", "", "", "", "", info, "si")


				 info = sImagenP1
				 sImagenP = "&#0232;"
				 sFuncAnc = ""
                 sImageOF =info

                  url = "'ORELRIESGOS.asp?SCODINDI=" &rs(0) & "' ,'mywindow',',scrollbars=1;resizable=0;status=0,toolbar=0,width=900,height=700'"




                  IF CINT("0" & RS("REL"))> 0 THEN
                     CELDAINFO = CELDAINFO & ofv.gencelda("", "", "CENTER onclick="&CHR(34)&"window.open("&url&");"&CHR(34) & " style='cursor:pointer;'", "", "", "", "", sImageOF, "si")
                  ELSE
                     CELDAINFO =CELDAINFO & ofv.gencelda("", "", "", "", "", "", "", "&nbsp;", "si")
                  END IF


                  info2 = sImagenI '"<B>USUARIOS RESPONSABLES</b>"
                  sImagenP = "&#0232;"
                  sFuncAnc = ""
                  SImageOF = "<font  size=1 color='#00AA99'  >" & INFO2 & "</font>"
                  url2 = "'ORELUSUARIOS.asp?SCODINDI=" &rs(0) & "' ,'mywindow',',scrollbars=1;resizable=0;status=0,toolbar=0,width=600,height=400'"

                  CELDAINFO = CELDAINFO & ofv.gencelda("", "", "CENTER onclick="&CHR(34)&"window.open("&url2&");"&CHR(34) & " style='cursor:pointer;'", "", "", "", "", sImageOF, "si")
                  CELDAINFO = CELDAINFO & CELDAINFOSEM & ofv.gencelda("", "", "CENTER", "", "", "", "", "&nbsp;", "si")
                  IF SVAUMBRAL="SI" OR SORUMBRAL="TODOS" THEN
                   FILAinfo = FILAinfo & ofv.genrow("aa bgcolor=" & xTrColor, "", "", "", "", "", "", CELDAINFO, "si")
                  END IF


                  CALL ofv.cerrarconsulta("rsum")
                  CALL ofv.cerrarconsulta(vars)


               END IF
            CASE ELSE

               IF NOT vars.EOF THEN
                  VALOR = REPLACE(vars("VALOR")&"", ",", ".")
                  FECHACT = vars("FECHACT")
                  ultresul=cdbl("0" & vars("ultresult"))

               ELSE
                  VALOR = 0
                  FECHACT = 0
                  ultresul=0
               END IF


               IF CDBL("0" & VALOR) = 0 AND CLNG("0" & FECHACT) = 0 THEN
                  descripcionUM = "NO EXISTEN VARIABLES"
                  sImageva = "<font face='corflaticon' size=6 color='#ff0000'>"&sImagevaC&"</font>"
               ELSE

                  Strsql = "Select * From orumbrales  WHERE VALOR1 <=0"&VALOR& " AND VALOR2>=0"& VALOR & " AND IDINDICADOR="&rs(0) 'id

                  SVAUMBRALFUERA="NO"
                  SET rsum = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

                  IF rsum.EOF THEN
                     '
                     sImageva = "<font face='corflaticon' size=6 color='#ff0000'>"&sImagevaC&"</font>"
                     descripcionUM = "FUERA DE RANGOS"
                     SVAUMBRALFUERA="SI"
                  ELSE

                     sImageva = "<font face='corflaticon' size=6 color='" & rsum("color") & "'>"&sImagevaC&"</font>"
                     descripcionUM = rsum("descripcion")

                     IF CSTR(RSUM("ORDEN")&"")= CSTR(SORUMBRAL&"") THEN  SVAUMBRAL="SI"

                  END IF
               END IF


               CELDAINFOSEM = ofv.gencelda("", "ut   align=CENTER" , "", "", "", "", "", sImageva, "si")

               info = "<u><b>CODIGO:" & RS(0)&"&nbsp;&nbsp;"& rs("descripcion") & "</b></u>"
               CELDAINFO =  ofv.gencelda("", "ut ALIGN=LEFT " , "", "", "", "", "", info, "si")

               'info = "<b>" & sinputVal &"--" & GTFechadorSH(FECHACT)& "</b>"

               info = "<b>" & sinputVal &"</b>"
               CELDAINFO = CELDAINFO & ofv.gencelda("", "ut  ALIGN=CENTER" , "", "", "", "", "", info, "si")




               'info = "UMBRAL :&nbsp;<b>" & descripcionUM &"</b>&nbsp;&nbsp;&nbsp; VALOR:&nbsp;<B>"& VALOR & "</B><br>"
               info =  descripcionUM
               CELDAINFO = CELDAINFO & ofv.gencelda("", "ut  ALIGN=CENTER" , "", "", "", "", "", info, "si")
               info =  VALOR
               CELDAINFO = CELDAINFO & ofv.gencelda("", "ut  ALIGN=RIGHT" , "", "", "", "", "", info, "si")


               info = sImagenP1 '"<B>VER RIESGOS ASOCIADOS</b>"
               sImagenP = "&#0232;"
               sFuncAnc = ""
               sImageOF =info ' "<font  size=1 color='#FF0000'  >" & INFO & "</font>"
               url = "'ORELRIESGOS.asp?SCODINDI=" &rs(0) & "' ,'mywindow',',scrollbars=1;resizable=0;status=0,toolbar=0,width=900,height=700'"




               IF CINT("0" & RS("REL"))> 0 THEN
                  CELDAINFO = CELDAINFO & ofv.gencelda("", "", "CENTER onclick="&CHR(34)&"window.open("&url&");"&CHR(34) & " style='cursor:pointer;'", "", "", "", "", sImageOF, "si")
               ELSE
                  CELDAINFO = CELDAINFO & ofv.gencelda("", "", "", "", "", "", "", "&nbsp;", "si")
               END IF



               info2 = sImagenI'"<B>USUARIOS RESPONSABLES</b>"
               sImagenP = "&#0232;"
               sFuncAnc = ""
               SImageOF = "<font  size=1 color='#00AA99'  >" & INFO2 & "</font>"
               url2 = "'ORELUSUARIOS.asp?SCODINDI=" &rs(0) & "' ,'mywindow',',scrollbars=1;resizable=0;status=0,toolbar=0,width=600,height=400'"

               CELDAINFO = CELDAINFO & ofv.gencelda("", "", "CENTER onclick="&CHR(34)&"window.open("&url2&");"&CHR(34) & " style='cursor:pointer;'", "", "", "", "", sImageOF, "si")
               url2 = "'orcargavariables.asp?Volver=NO&opcion=VERINFO&id="& RS(0) & "' ,'mywindow',',scrollbars=1;resizable=1;status=0,toolbar=0,width=800,height=500'"
			   select case cdbl(valor)
			    case valor > ultresul
			    CELDAINFOR= ofv.gencelda("", "", "CENTER onclick="&CHR(34)&"window.open("&url2&");"&CHR(34) & " style='cursor:pointer;'", "", "", "", "",  flechaup , "si")
			   case valor < ultresul
				 CELDAINFOR= ofv.gencelda("", "", "CENTER onclick="&CHR(34)&"window.open("&url2&");"&CHR(34) & " style='cursor:pointer;'", "", "", "", "",  flechadown , "si")
			    case else
				CELDAINFOR= ofv.gencelda("", "", "CENTER onclick="&CHR(34)&"window.open("&url2&");"&CHR(34) & " style='cursor:pointer;'", "", "", "", "",flechaequal , "si")
			 end select

                CELDAINFO = CELDAINFO & CELDAINFOSEM & CELDAINFOR
            '   RESPONSE.WRITE SORUMBRAL & "---" & SVAUMBRALFUERA
               IF SORUMBRAL="FUERARANGOS" AND SVAUMBRALFUERA="SI" THEN
                 FILAinfo = FILAinfo & ofv.genrow("aa bgcolor=" & xTrColor , "", "", "", "", "", "", CELDAINFO, "si")
               ELSE
               IF SVAUMBRAL="SI" OR SORUMBRAL="TODOS" THEN
                 FILAinfo = FILAinfo & ofv.genrow("aa bgcolor=" & xTrColor , "", "", "", "", "", "", CELDAINFO, "si")
               END IF
               END IF

               CALL ofv.cerrarconsulta("rsum")
               CALL ofv.cerrarconsulta(vars)

               posi = posi + 1
         END SELECT
         RS.MOVENEXT


      LOOP
   END IF

  ' IF fila = "" OR celda<>"" THEN fila = fila & ofv.genrow("", "", "", "", "90", "", "", Celda, "si")

   tabla2 = ofv.GenTabla("", "aaa  ", "", "95% ", "", "0", "", "1", "1", "", "", FILAinfo, "si")

   sHTML = sHTML & "<CENTER>" & tabla1 & tabla2 & "</center>"




   'Botonera
   Agregar = ""
   Buscar = ""
   Celda = ""
   sAccion = "ORBUSCAINDI.ASP"
   Celda = ofv.BotonVolver(sAccion)

   sHTML = sHTML & ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "", Agregar, Buscar, celda)

   CALL ofv.cerrarconsulta(rs)
   CALL ofv.cerrarconn(cn)

   sHTML = sHTML & "</body></html>"
END IF

FUNCTION validarFecuencia2(Idindicador, fechaultact, dias)


   sinput = "SI"
   DIAS = "0" & DIAS
   FEC = CDATE( dias + CLNG("0" & fechaultact) + CLNG(CDATE("01-01-1900") ))

   IF DATE()>= fec THEN
      SInput = "SI"
   ELSE
      SINPUT = "NO"
   END IF
   validarFecuencia2 = sinput '& fec & "--" & date()
END FUNCTION

Response.WRITE sHTML

%>
