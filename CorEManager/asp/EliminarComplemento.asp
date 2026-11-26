<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo = request.querystring("Tipo")
sNombre = request.querystring("Nombre")
sID = request.querystring("ID")
sEliminar = request.querystring("Eliminar")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
if sEliminar = "NSNC" then session("volver2") = volver

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "Select * From Complementos Where codcomp = " & sID
set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
if not rs.eof then sComp = rs(3)
call ofv.cerrarconsulta(rs)

xsen = "N"

  Strsql = "SELECT codobjeto From Objestdet where detimagen = '" & sID & "' "
  Strsql = Strsql & " and dettipo = 'J'  "
 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then xsen = "S"

  strsql = "Select * From Objaccion Where codcomp = " & sID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then xsen = "S"


if xsen = "S" then
  sMensaje = "No puede eliminar " & sComp & " porque esta siendo utilizado "
  sMensaje = sMensaje & "en algun esquema normativo "
  dato = "../Mensajes.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver=" & replace(replace(volver,"&","_"),"../","")
  dato = dato & "&Tipo=" & sTipo & "&Nombre=" & sNombre & "&Botones=1" & chr(13)
'  response.write dato & "<br>"
  response.redirect dato
else
  if sEliminar = "NSNC" then
    sMensaje = "Esta seguro que desea eliminar " & sComp
    dato = "../Mensajes.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
    dato = dato & "asp/EliminarComplemento.asp?ID=" & sID & "&Tipo=" & sTipo & "&Nombre="
    dato = dato & sNombre & "&Botones=2" & chr(13)
    response.redirect dato
  elseif sEliminar = "SI" then
    call ofv.cerrarconsulta(rs)
    strsql = "Delete From Compimag Where codcomp = " & sID
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
    strsql = "Delete From Compprty Where codcomp = " & sID
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
    strsql = "Delete From Compdet Where codcomp = " & sID
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
    strsql = "Delete From Compestr Where codcomp = " & sID
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
    strsql = "Delete From Compcompuesto Where compasoc = " & sID
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
    strsql = "Delete From Complementos Where codcomp = " & sID
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
  end if
end if

Volver = session("volver2")
''''volver = "../" & volver & "?Tipo=" & sTipo & "&Nombre=" & sNombre
volver = volver & "&Tipo=" & sTipo & "&Nombre=" & sNombre

'''response.write volver & "<br>"

set errores=cn.errors
if errores.count = 0 then
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>