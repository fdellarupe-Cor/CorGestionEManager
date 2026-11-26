<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
sAcc = request.querystring("acc")
sCalif = request.querystring("calif")
sCodSec = request.querystring("codsector")

set cn = ofv.conectar(ofv.strconn2)

if UCASE(sAcc) = "CARGA" then

	sql = "insert into objsector (codobjeto,tipobjeto,calificador,codsector) "
	sql = sql & " values ('" & session("NCliente") & "','" & session("TCliente") & "','" & sCalif & "','" & sCodSec & "')"

elseif UCASE(sAcc) = "ELIM" then

	sql = "delete from objsector where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "' and codsector=" & sCodSec

end if

'response.write sql
call ofv.crearconsultaEx(sql,cn,1,parametros)

%>