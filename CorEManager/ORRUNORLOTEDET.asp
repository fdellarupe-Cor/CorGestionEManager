<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 20
  
  
  set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar("SELECCIONADOS","_"," ")
    response.write  oPrtG.RepHeader2(cstr(""), cstr("600"))
	response.write corstyler()
  
  'CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  'vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  'sopcion  = request.querystring("opcion") & ""

  'sHTML = ofv.FormHeader(session("FormN"))

	'response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
	'response.write "<tr>"
	'response.write "<td width=100% >"
	'response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
	'response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
	'response.write "<tr>"
	'response.write "<td width=10% >&nbsp;</td>"
	'response.write "<td width=90% >&nbsp;</td>"
	'response.write "</tr></table></td></tr></table>"


  Formulario = ""

  LOTE = REQUEST.QUERYSTRING("LOTE")

  if lote<> "" then
     session("slote")=LOTE
  else
     lote=session("slote")
  end if


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar




'
  '  'Encabezados


  set cn = ofv.conectar(ofv.strconn2)


  STRSQL = " SELECT a.objid  FROM ORREGINFOBCRADET A   "
  STRSQL = STRSQL & " WHERE a.CODLOTE='" &  LOTE & "' "
  ' RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
      Celda = ""

       Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","","NO EXISTEN REGISTROS SELECCIONADOS","si")
       Fila  = ofv.GenRow("","","","","","","",Celda,"si")
       sHTML =  shtml &  "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

       FILA=""
     'Response.Write shtml

  else


     '' STRSQL = " SELECT a.objid as idlote , A.TRIMESTRE AS TRIMESTREZ ,A.ANO ,B.* FROM ORREGINFOBCRA A "
     '' STRSQL = STRSQL & "inner JOIN ORREGINFOBCRADET B  ON (A.CODLOTE=B.CODLOTE) "
     '' STRSQL = STRSQL & " WHERE a.CODLOTE='" &  LOTE &"' ORDER BY B.OBJIDORTRACER"



      STRSQL = " SELECT a.objid as idlote , A.TRIMESTRE AS TRIMESTREZ ,A.ANO ,B.* FROM ORREGINFOBCRA A "
      STRSQL = STRSQL & "inner JOIN ORREGINFOBCRADET B  ON (A.CODLOTE=B.CODLOTE) "
      STRSQL = STRSQL & " WHERE a.CODLOTE='" &  LOTE & "' ORDER BY B.YANIO, B.TRIMESTRE, B.ESTADO, B.OBJIDORTRACER"
  ' RESPONSE.WRITE STRSQL
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rs.eof then

         sInput = "DETALLE DEL LOTE :&nbsp;<b>"& rs("idlote")&"</b> &nbsp;&nbsp;&nbsp;&nbsp;TRIMESTRE :&nbsp;<b>"& rs("TRIMESTREZ") & "</b>&nbsp;&nbsp;&nbsp;AÑO :&nbsp;<b>"& rs("ANO")& "</b>"
         Celda = ofv.gencelda("","TT ","","90%","","","",sInput,"si")

         Fila = ofv.genrow("","","","","","","",Celda,"si")

         Formulario = Fila
         sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"

         Fila=""



         xTrColor = "#FFFFFF"

         posI = 0
         DIM TOTAL
         DIM TOTAL2

         WANTAA = 0
         WANTTT = 0
         WANTTY = "-"

      DO WHILE Not rs.EOF



         WACTAA = RS("YANIO")
         WACTTT = RS("TRIMESTRE")
         WACTTY = RS("ESTADO")

    IF (WACTAA <> WANTAA) OR (WACTTT <> WANTTT)  OR (WACTTY <> WANTTY) THEN

        IF WANTTT = 0 THEN




        ELSE


           Celda = ofv.gencelda("","CC COLSPAN=6 STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","","<B>TOTAL DEL ARCHIVO</B>","si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","","","","",wwTOTRR,"si")
           Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","","<B>MONTO</B>","si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","","","","",TOTALR,"si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","","","","",TOTALR2,"si")
           Fila  = FILA & Ofv.genrow("","","","","","","",Celda,"si")

           Celda = ofv.gencelda("","UT COLSPAN=9  " ,"","","","","","&nbsp;","si")
           Fila  = FILA & Ofv.genrow("","","","","","","",Celda,"si")




        END IF

        Fila  = FILA & ofv.Encabezados("EVENTO","IMPORTE","CARACTER","EV.BCRA","SEC.BCRA","TRIMESTRE","AÑO","TIPO","IMP.BCRA")


        wwTOTRR = 0
        TOTALR  = 0
        TOTALR2  = 0

        WANTAA  = RS("YANIO")
        WANTTT  = RS("TRIMESTRE")
        WANTTY  = RS("ESTADO")




    END IF





         sForm = "form" & posi


      'CODIGO


         wwTOTRR = wwTOTRR + 1

         WWWMONTOEV = CDBL(RS("MONTO"))

         IF (RS("CARACTER") & "") = "06" OR (RS("CARACTER") & "") = "07" THEN
             WWWMONTOEV2 = round(CDBL(RS("MONTO")) * -1 ,0) 
         ELSE
             WWWMONTOEV2 = round(CDBL(RS("MONTO")) , 0)
         END IF




         TOTAL    = TOTAL    + WWWMONTOEV
         TOTAL2   = TOTAL2   + WWWMONTOEV2
         TOTALR   = TOTALR   + WWWMONTOEV
         TOTALR2  = TOTALR2  + WWWMONTOEV2



         Celda = ofv.gencelda("","VTI COLSPAN=2 STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","10%","","","",RS("OBJIDORTRACER") & "","si")
     ''''    Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","20%","","","","$" & replace(RS("MONTO")&"",",","."),"si")
         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","20%","","","",replace(WWWMONTOEV & "",",","."),"si")
         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","15%","","","", RS("CARACTER") & "","si")
         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","15%","","","", RS("EVENTORAIZ") & "","si")

         WWWSECUEN =  RS("SECUENCIA") & ""

         IF WWWSECUEN = "-1" THEN WWWSECUEN = "REP."

         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","5%","","","", WWWSECUEN,"si")


         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","5%","","","", RS("TRIMESTRE") & "","si")
         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","5%","","","", RS("YANIO") & "","si")

         WWKESTADO = "NORMAL"


         SELECT CASE TRIM(RS("ESTADO") & "")

                CASE "J"
                      WWKESTADO = "<FONT COLOR='#993333'><B>AJUSTE</B></FONT>"

                CASE "R"
                      WWKESTADO = "<FONT COLOR='#000088'><B>RECTIFICATIVA</B></FONT>"

         END SELECT

         WWKESTADO = WWKESTADO


         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","15%","","","", WWKESTADO & "","si")
         Celda = CELDA & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","20%","","","",replace(WWWMONTOEV2 & "",",","."),"si")


         Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")


         if xTrColor = "#FFFFFF" then
            xTrColor = "#DFDFDF"
         ELSE
            xTrColor = "#FFFFFF"
         END IF


        RS.MOVENEXT
        posI=posI+1
      LOOP


      IF POSI > 0 THEN

           Celda = ofv.gencelda("","CC COLSPAN=6 STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","0%","","","","<B>TOTAL DEL ARCHIVO</B>","si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","0%","","","",wwTOTRR,"si")
           Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","0%","","","","<B>MONTO</B>","si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","0%","","","",TOTALR,"si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","0%","","","",TOTALR2,"si")
           Fila  = FILA & Ofv.genrow("","","","","","","",Celda,"si")

           Celda = ofv.gencelda("","UT COLSPAN=9  " ,"","","","","","&nbsp;","si")
           Fila  = FILA & Ofv.genrow("","","","","","","",Celda,"si")


      END IF




     end if


           Celda = ofv.gencelda("","CC COLSPAN=6 STYLE='border-style:solid;border-width:1;border-color:#000000;' " ,"","0%","","","","<B>TOTAL GENERAL</B>","si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#000000;text-align:right' " ,"","0%","","","",POSI,"si")
           Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#000000;' " ,"","0%","","","","<B>MONTO</B>","si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#000000;text-align:right' " ,"","0%","","","",TOTAL,"si")
           Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#000000;text-align:right' " ,"","0%","","","",TOTAL2,"si")
           Fila  = FILA & Ofv.genrow("","","","","","","",Celda,"si")


   sHTML = sHTML  & "<BR>" & ofv.GenTabla("","","","60% ","","0","","0","0","","",FILA,"si")

  end if


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


Response.Write sHTML

end if

%>