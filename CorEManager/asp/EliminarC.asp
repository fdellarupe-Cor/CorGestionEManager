<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

nCodigo = request.querystring("Codcomp")
sTipo = request.querystring("Tipocomp")
sID = request.querystring("ID")
Tabla = request.querystring("Tabla")
sEliminar = request.querystring("Eliminar")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
if sEliminar = "NSNC" then session("volver2")=volver

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "Select * From " & Tabla & " Where objid = " & sID
set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
if not rs.eof then
  if Tabla = "Compestr" or Tabla = "Compimag" then
    sComp = rs(3)
  else
    sComp = rs(4)
  end if
end if
call ofv.cerrarconsulta(rs)

Volver = session("volver2")
volver = volver & "&Tipocomp=" & sTipo & "&Codcomp=" & nCodigo

if sEliminar = "NSNC" then
  sMensaje = "Esta seguro que desea eliminar " & sComp
  dato = "../Mensajes2.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
  dato = dato & "asp/EliminarC.asp?ID=" & sID & "&Tipocomp=" & sTipo & "&Codcomp="
  dato = dato & nCodigo & "&Botones=2" & "&Tabla=" & Tabla & chr(13)
  response.redirect dato
elseif sEliminar = "SI" then
  if Tabla = "Compestr"  then
     strsql = "Delete From Compdet Where codestr = " & sID
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
  end if
  strsql = "Delete From " & Tabla & " Where objid = " & sID
  call ofv.crearconsultaEx(StrSql,cn,1,volver)

  set errores = cn.errors
  if errores.count = 0 then
    response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
  end if
else
  response.redirect volver
end if

call ofv.cerrarconn(cn)

%>