<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo     = request.querystring("Tipo")
sNombre   = request.querystring("Nombre")
sID       = request.querystring("ID")
sEliminar = request.querystring("Eliminar")
Volver    = request.querystring("Volver")
Volver    = ofv.convertircar(Volver,"_","&")

if sEliminar = "NSNC" then session("volver2") = volver

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "Select * From Complementos Where codcomp = " & sID
st=strsql
set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
if  not rs.eof then sComp = rs(3)
call ofv.cerrarconsulta(rs)
Smsg=""
'strsql = "Select * From Objlnk Where ltipo='KL' and codobjetoL = " & sID
strsql="select OBJID from objaccion where codcomp  in( SELECT objid FROM OBJLNK a WHERE a.LTIPO='KL' AND a.CODOBJETOL = "& sID &" )"
STRSQL =STRSQL & " UNION "
STRSQL = STRSQL & "select OBJID from objestdet where  detimagen  in( SELECT convert(varchar(100),objid) FROM OBJLNK a WHERE a.LTIPO='KL' AND a.CODOBJETOL = "&SID&" )"
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)





if not rs.eof then
  sMensaje = "<b>No puede eliminar <u>" & sComp & "</u> porque esta siendo utilizado " 	'& st
  sMensaje = sMensaje & "en algun Documento ...!!!</b>"
  sMensaje = Ucase(sMensaje)
  dato = "../Mensajes.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver=" & volver
  dato = dato & "&Tipo=" & sTipo & "&Nombre=" & sNombre & "&Botones=1" & chr(13)
  response.redirect dato
else
  if sEliminar = "NSNC" then
    sMensaje = "<b>Confirma eliminar <u>" & sComp & "</u> ...?</b> " '& st
    sMensaje = Ucase(sMensaje)
    dato = "../Mensajes.asp?msg=" & ofv.convertircar(sMensaje," ","_")
    dato = dato & "&Tipo=" & sTipo & "&Nombre="
    dato = dato & sNombre & "&Botones=2" & "&volver=asp/ElComp.asp?ID=" & sID & chr(13)
    response.redirect dato
  ' RESPONSE.WRITE DATO & "<BR>" & SESSION("VOLVER2") & "<BR>"
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
volver = "../" & volver & "&Tipo=" & sTipo & "&Nombre=" & sNombre

 '  RESPONSE.WRITE DATO & "<BR>" & SESSION("VOLVER2") & "<BR>" & volver & "<BR>"

set errores=cn.errors
if  errores.count = 0 then
    response.redirect volver
else
    response.write ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>