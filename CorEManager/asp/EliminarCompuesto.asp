<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nCodigo = request.querystring("Codcomp")
sTipo = request.querystring("Tipocomp")
sID = request.querystring("ID")
Tabla = request.querystring("Tabla")
sEliminar = request.querystring("Eliminar")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
if sEliminar = "NSNC" then session("volver2") = volver

Volver = session("volver2")
volver = volver & "?Tipocomp=" & sTipo & "&Codcomp=" & nCodigo

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "Select * From Cscompuesto Where id = " & sID
set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
if not rs.eof then sComp = rs(7)
call ofv.cerrarconsulta(rs)

if sEliminar = "NSNC" then
  sMensaje = "Esta seguro que desea eliminar " & sComp
'  sVolver = "EliminarCompuesto.asp?ID=" & sID & "&Tipocomp=" & sTipo & "&Codcomp=" & nCodigo
'  sVolver = sVolver & "&Tabla=" & Tabla
'  sHTML = ofv.Mensajes(sMensaje,sVolver,"2")
'  response.write sHTML
  dato = "../Mensajes2.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
  dato = dato & "asp/EliminarCompuesto.asp?ID=" & sID & "&Tipocomp=" & sTipo & "&Codcomp="
  dato = dato & nCodigo & "&Botones=2" & "&Tabla=" & Tabla & chr(13)
  response.redirect dato
elseif sEliminar = "SI" then
  strsql = "Delete From Compprty Where codcomp = " & sID
  call ofv.crearconsultaEx(StrSql,cn,1,volver)
  strsql= "Delete From " & Tabla & " Where id = " & sID
  call ofv.crearconsultaEx(StrSql,cn,1,volver)
else
  response.redirect volver
end if



set errores = cn.errors
if errores.count = 0 then
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>