<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

ID = request.querystring("ID")
t = request.querystring("t")
sMenu = request.querystring("Menu")
Nombre = ofv.convertircar(request.querystring("Nombre"),"_"," ")
sAplicacion = request.querystring("Aplicacion")
volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
volver = volver & "&ID=" & ID & "&t=" & t & "&Menu=" & sMenu & "&Nombre=" & Nombre
volver = volver & "&Aplicacion=" & sAplicacion

set cn = ofv.conectar(ofv.strconn0)

    strsql = "SELECT *  From AtrMUsuarios Where Usuario = '" & ID & "'"
    'RESPONSE.WRITE STRSQL
    SET RSATRM=ofv.crearconsultaEx(StrSql,cn,1,volver)
    IF NOT RSATRM.EOF THEN
      XID=RSATRM(0)
      COL="ID"
    ELSE
      XID=ID
      COL="USUARIO"
    END IF
    CALL OFV.CERRARCONSULTA(RSATRM)

    strsql = "Delete  From AtrMUsuarios Where Usuario = '" & ID & "'"

    call ofv.crearconsultaEx(StrSql,cn,1,volver)
    xxxr=GENLOGEXTDATA("DELETE","AtrMUsuarios",COL,XID,"SKR")

    response.redirect volver

call ofv.cerrarconn(cn)

%>