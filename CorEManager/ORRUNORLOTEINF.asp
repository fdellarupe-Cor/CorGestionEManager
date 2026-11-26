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


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sopcion  = request.querystring("opcion") & ""

 ' sHTML = ofv.FormHeader(session("FormN"))



  Formulario = ""

 LOTE=REQUEST.QUERYSTRING("LOTE")


	set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
	Set oPrtG.Ofv = Ofv
	oPrtG.RCliLogo = session("clilogo")
	oPrtG.RCliNomb = ofv.convertircar("TOTALES","_"," ")
    response.write  oPrtG.RepHeader2(cstr(""), cstr("600"))
	response.write corstyler()    'response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
   'response.write "<tr>"
   'response.write "<td width=100% >"
   ' response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
   'response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"

    'response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
    'response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
    'response.write "</tr></table>"

    'Encabezados



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)





'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LOTE NUEVO

  WKSETTABLA = "ORREGINFOBCRADET"

  WWWTOTREG = 0
  WWWTOTlei = 0
  WWWTOTsel = 0
  WWWTOTerr = 0
  WWWTOTRAI = 0
  WWWTOTREP = 0
  WWWTOTENC = 0
  WWWTOTREPA = 0

  WWWTOTREGm = 0
  WWWTOTleim = 0
  WWWTOTselm = 0
  WWWTOTerrm = 0
  WWWTOTRAIM = 0
  WWWTOTREPM = 0
  WWWTOTENCM = 0

  STRSQL = " SELECT  MAX(EVENTORAIZ), MAX(SECUENCIA), SUM(MONTO) , SUM(1)"
  STRSQL = STRSQL & " FROM ORREGINFOBCRADET B "
  STRSQL = STRSQL & " WHERE B.CODLOTE='" &  LOTE &"'"
  STRSQL = STRSQL & " GROUP BY EVENTORAIZ, SECUENCIA"
'  response.write strsql & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  DO WHILE NOT RS.EOF
     WWWTOTREG = WWWTOTREG + 1

     IF (RS(1) & "") = "0" THEN
         WWWTOTRAI  = WWWTOTRAI + 1
         WWWTOTRAIM = WWWTOTRAIM + CDBL(rs(2))
     END IF

     IF (RS(1) & "") = "-1" THEN
         WWWTOTREP  = WWWTOTREP + CLNG(rs(3))
         WWWTOTREPA = WWWTOTREPA + 1
         WWWTOTREPM = WWWTOTREPM + CDBL(rs(2))
     END IF

     IF (RS(1) & "") > "0" THEN
         WWWTOTENC  = WWWTOTENC + 1
         WWWTOTENCM = WWWTOTENCM + CDBL(rs(2))
     END IF


   RS.MOVENEXT
  LOOP


  STRSQL = " SELECT  SUM(1), SUM(MONTO) "
  STRSQL = STRSQL & " FROM ORREGINFOBCRADET B "
  STRSQL = STRSQL & " WHERE B.CODLOTE='" &  LOTE &"'"

'  response.write strsql & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  DO WHILE NOT RS.EOF
      WWWTOTsel  = rs(0)
      WWWTOTselm = CDBL(rs(1))
   RS.MOVENEXT
  LOOP

'''ORRIERRLOTE

  STRSQL = " SELECT  SUM(1), SUM(D.TRCIMPBRUTO) from ("
  STRSQL = STRSQL & " select distinct B.nroevento, B.CODLOTE from ORRIERRLOTE B "
  STRSQL = STRSQL & " WHERE B.CODLOTE='" &  LOTE &"' ) C "
  STRSQL = STRSQL & " INNER JOIN ORTRACER D ON (C.NROEVENTO = D.OBJID)  "

'  response.write strsql & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  DO WHILE NOT RS.EOF
      WWWTOTerr  = rs(0)
      WWWTOTerrM = CDBL(rs(1))
   RS.MOVENEXT
  LOOP


  STRSQL = " SELECT MAX(A.TOTAL), MAX(A.totalMontoTri) , SUM(MONTO), MAX(EVENTORAIZ), MAX(TOTALINFORMAR), MAX(TOTALNOINFORMAR), "
  STRSQL = STRSQL & " max(totrep) AS totrep, max(totenca) AS totenca, max(totdifrep) AS totdifrep, max(a.trimestre), max(a.ano),"
  STRSQL = STRSQL & " MAX(MONTOTOTALNOINF) AS MONTOTOTALNOINF , max(a.objid) as idlote "
  STRSQL = STRSQL & " FROM ORREGINFOBCRA A "
  STRSQL = STRSQL & " INNER JOIN ORREGINFOBCRADET B  ON (A.CODLOTE=B.CODLOTE) "
  STRSQL = STRSQL & " WHERE B.CODLOTE='" &  LOTE &"'"
'  response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  formulario= "<tr> <td class=tt colspan=3 >LOTE : "& rs("idlote") &" TRIMESTRE :&nbsp;"& rs(9) & " AÑO :&nbsp;"& rs(10 )&"</td></tr>"


  Formulario=formulario & chr(13)&" <TR ><TD  class=tt width=60% >item</TD> <TD  class=tt width=20% >cantidad</TD><TD  class=tt width=20% >Importe</TD></tr>" & chr(13)


  posI = 0
  IF rs("idlote")&"" <> "" THEN
    sForm = "form" & posi


    'CODIGO

    Celda =  ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","",""," REGISTROS DEL TRIMESTRE","si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","", WWWTOTsel + WWWTOTerr,"si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","",FormatNumber((WWWTOTerrM + WWWTOTselm),2),"si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","","REGISTROS SELECCIONADOS","si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;;text-align:right;' " ,"","","","","", RS(4),"si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","", FormatNumber(RS(2),2),"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ofv.gencelda("","UTI STYLE='border-style:sOlid;border-width:1;border-color:#888888;' " ,"","","","","","REGISTROS NO SELECCIONADOS","si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;;text-align:right;' " ,"","","","","", RS(5),"si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;;text-align:right;' " ,"","","","","", FormatNumber(RS("MONTOTOTALNOINF"),2),"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ofv.gencelda("","UTI STYLE='border-style:sOlid;border-width:1;border-color:#888888;' " ,"","","","","","EVENTOS REPETITIVOS","si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;text-align:right;' " ,"","","","","", WWWTOTREP,"si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","", "&nbsp;","si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ofv.gencelda("","UT COLSPAN=3 STYLE='border-style:sOlid;border-width:2;border-color:#000000;background-color:#bbbbbb;font-size:2pt;' " ,"","","","","","&nbsp;","si")
    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","VTI STYLE='border-style:sOlid;border-width:1;border-color:#000088;' " ,"","","","","","EVENTOS RAIZ REPETITIVOS AGRUPADOS","si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#000088;text-align:right;' " ,"","","","","", WWWTOTREPA,"si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#000088;text-align:right;' " ,"","","","","",FormatNumber(WWWTOTREPM,0) ,"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ofv.gencelda("","VTI STYLE='border-style:sOlid;border-width:1;border-color:#000088;' " ,"","5%","","","","EVENTOS ENCADENADOS","si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#000088;text-align:right;' " ,"","","","","", WWWTOTENC,"si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#000088;text-align:right;' " ,"","","","","", FormatNumber(WWWTOTENCM,0),"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ofv.gencelda("","VTI STYLE='border-style:sOlid;border-width:1;border-color:#000088;' " ,"","5%","","","","EVENTOS RAIZ E INDEPENDIENTES","si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#000088;text-align:right;' " ,"","","","","", WWWTOTRAI,"si")
    Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#000088;text-align:right;' " ,"","","","","", FormatNumber(WWWTOTRAIM,0),"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ofv.gencelda("","UT COLSPAN=3 STYLE='border-style:sOlid;border-width:2;border-color:#000000;background-color:#bbbbbb;font-size:2pt;' " ,"","","","","","&nbsp;","si")
    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","VTI STYLE='border-style:sOlid;border-width:1;border-color:#880000;' " ,"","","","","","REGISTROS A ENVIAR","si")
    Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#880000;;text-align:right;' " ,"","","","","", WWWTOTREG,"si")
    Celda = CELDA & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#880000;;text-align:right;' " ,"","","","","", FormatNumber(WWWTOTRAIM + WWWTOTENCM + WWWTOTREPM,0),"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")



  else
     STRSQL=" SELECT * from ORREGINFOBCRA WHERE CODLOTE='" &  LOTE &"'"
	' RESPONSE.WRITE STRSQL
	CALL OFV.CERRARCONSULTA(RS)
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then

     Celda =  ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","",""," REGISTROS DEL TRIMESTRE","si")
	 Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","", RS("total"),"si")
	 Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","",FormatNumber(RS("totalmontotri"),2),"si")
	 Fila = ofv.genrow("","","","","","","",Celda,"si")


     Celda = ofv.gencelda("","CC STYLE='border-style:sOlid;border-width:1;border-color:#888888;' " ,"","","","","","REGISTROS NO SELECCIONADOS","si")

     Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;;text-align:right;' " ,"","","","","", RS("total"),"si")
     Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;;text-align:right;' " ,"","","","","", FormatNumber(RS("MONTOTOTALNOINF"),2),"si")
	 Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")
     Celda = ofv.gencelda("","CC STYLE='border-style:sOlid;border-width:1;border-color:#888888;' " ,"","","","","","REGISTROS A ENVIAR","si")
	 Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;;text-align:right;' " ,"","5%","","","", "0","si")
	 Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;;text-align:right;' " ,"","5%","","","", "&nbsp;","si")
	 Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")
	  Celda = ofv.gencelda("","CC STYLE='border-style:sOlid;border-width:1;border-color:#888888;' " ,"","","","","","EVENTOS REPETITIVOS","si")
	  Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;text-align:right;' " ,"","","","","", RS("totrep"),"si")
	 Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","", "&nbsp;","si")
	  Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")



	     Celda = ofv.gencelda("","CC STYLE='border-style:sOlid;border-width:1;border-color:#888888;' " ,"","","","","","EVENTOS ENCADENADOS","si")
	 	 Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","", RS("totenca"),"si")
	 	 Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","", "&nbsp;","si")
	 	 Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

        Celda = ofv.gencelda("","CC STYLE='border-style:sOlid;border-width:1;border-color:#888888;' " ,"","","","","","EVENTOS RAIZ REPETITIVOS AGRUPADOS","si")
		Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","", RS("totdifrep"),"si")
		Celda = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","", "&nbsp;","si")
		Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    end if

  END IF

  sHTML = sHTML  & ofv.GenTabla("","","","80% ","","0","","0","0","","",formulario & FILA,"si") & "</body></html>"
end if

Response.Write sHTML

CALL OFV.CERRARCONSULTA(RS)
call ofv.cerrarconn(cn)

%>