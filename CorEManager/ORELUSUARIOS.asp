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


 SHTML=SHTML & USUARIOSRESP(SCODINDI)  &  "</body></html>"
 call ofv.cerrarconn(cn)
Response.Write sHTML
END IF


FUNCTION USUARIOSRESP(SCODINDI)
DIM RS
DIM CELDA
DIM SHTML
DIM Formulario
DIM sForm

  Fila=""
  WKSETTABLA = "ORUSRLST"

  Strsql = "Select *  From " & WKSETTABLA & " where indicador="& SCODINDI

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posI = 1
  FILA=""
  Celda =ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","40% ALIGN=CENTER","","","","<big>USUARIOS RESPONSABLES</BIG>","si")
  Fila = ofv.genrow("","","","","","","",CELDA,"si")
  Do While (Not rs.EOF)

    XXX=GETINFOUSER(REPLACE(RS("USRRESP")&"","USR-",""))
    Celda =ofv.gencelda("","UTI ","LEFT","","","","",session("SGU_nombre") ,"si")

    Fila = FILA & ofv.genrow("","","","","","","",CELDAX & Celda,"si")
    CALL OFV.CERRARCONSULTA(RSRIESGO)

    rs.Movenext
  Loop



    USUARIOSRESP = ofv.GenTabla("","","","98%% ","","0","","0","0","","",FILA,"si")

   CALL OFV.CERRARCONSULTA(rs)
END FUNCTION

%>