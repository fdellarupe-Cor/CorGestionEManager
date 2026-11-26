<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

 set cn = ofv.conectar(ofv.strconn2)

 SCODINDI=REQUEST.QUERYSTRING("SCODINDI")


 sHTML = ofv.FormHeader(session("FormN"))


 SHTML=SHTML & RIESGOSASOCIADOS(SCODINDI)  &  "</body></html>"
 call ofv.cerrarconn(cn)
Response.Write sHTML
END IF


FUNCTION RIESGOSASOCIADOS(SCODINDI)
DIM RS
DIM CELDA
DIM SHTML
DIM Formulario
DIM sForm

' Strsql = "Select * From CSRELPATHDESCRI  where COBJ
  Fila=""
  WKSETTABLA = "ORRELAINDIRIESGO"

  Strsql = "Select CODRIESGO  From " & WKSETTABLA & " where idindicador="& SCODINDI
 ' RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posI = 1
  FILA=""
  Do While (Not rs.EOF)


    RUTA=""
    STRSQL="SELECT * FROM CSORBUSQRIESGO WHERE CODRIESGO='" & RS(0) & "'"
    SET RSRIESGO=ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Celda =ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RSRIESGO("RIESGONOMBRE"),"si")
    Celda = CELDA & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RSRIESGO("DESCRIACT"),"si")
    Celda =CELDA & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RSRIESGO("PERIODO"),"si")
    if not RSRIESGO.eof then
    Strsql = "Select * From CSRELPATHDESCRI  where COBJ=" & RSRIESGO("CODOBJETO")
	SET RSRIESGOPATH=ofv.crearconsultaEx(StrSql,cn,1,parametros)
    DELIMX=""
	FOR I=1 TO 7
	  IF RSRIESGOPATH("EXPR" & I)&"" <> "" THEN
	    RUTA=RUTA & DELIMX & RSRIESGOPATH("EXPR" & I)
	    DELIMX="/"
	  END IF
    NEXT
    CELDAX=ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RUTA,"si")
    Fila = FILA & ofv.genrow("","","","","","","",CELDAX & Celda,"si")
    CALL OFV.CERRARCONSULTA(RSRIESGO)
    end if
    rs.Movenext
  Loop

  IF FILA <> "" THEN
    Celda = "<BR>" & ofv.gencelda("","VTI  COLSPAN=4 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>LISTA DE RIESGOS ASOCIADOS</BIG>","si")
    FILAC=ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    Celda = ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","40% ALIGN=CENTER","","","","<big>PROCESO/SUBPROCESOS</BIG>","si")
	Celda = CELDA & ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","35% ALIGN=CENTER","","","","<big>RIESGO</BIG>","si")
	Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15% ALIGN=CENTER","","","","<big>ACTIVIDAD</BIG>","si")
	Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","10% ALIGN=CENTER","","","","<big>PERIODO</BIG>","si")
    FilaC  = FILAC & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    RIESGOSASOCIADOS = ofv.GenTabla("","","","98%% ","","0","","0","0","","",FILAC & Fila,"si")
 ELSE
    RIESGOSASOCIADOS=""
 END IF
   CALL OFV.CERRARCONSULTA(rs)
END FUNCTION

%>