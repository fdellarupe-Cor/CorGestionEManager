<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

sID = request.querystring("ID")
sEliminar = request.querystring("Eliminar")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
if sEliminar = "NSNC" then session("volver2") = volver

set cn = ofv.conectar(ofv.strconn2)

strsql = "Select * From Prodserv Where codprodserv = " & sID
sMsg = ""
set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
if not rs.eof then
   tipoin = " tipobjmaster in ('999'"
   strsql = "Select tipobjeto From Subtema "
   set rsx = ofv.crearconsultaEx(strsql,cn,1,volver)
   do until rsx.eof
      tipoin = tipoin & ",'" & rsx(0) & "'"
   rsx.movenext
   loop
   tipoin = tipoin & ") "
   call ofv.cerrarconsulta(rsx)

  sObjeto = rs(1)

  strsql = "Select * From Relaobj Where codobjmaster = " & sID & " and "
  strsql = strsql & tipoin
  set rsx = ofv.crearconsultaEx(strsql,cn,1,volver)
  if not rsx.eof then sMsg = "No puede eliminar " & sObjeto & " porque tiene esquemas asociados"
  call ofv.cerrarconsulta(rsx)
end if
call ofv.cerrarconsulta(rs)

if sEliminar <> "NO" then
  if sMsg = "" then
    if sEliminar = "NSNC" then
      sMensaje = "Esta seguro que desea eliminar " & sObjeto
      dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
      dato = dato & "asp/EliminarSubtema.asp?ID=" & sID & "&Botones=2"
      response.redirect dato
    elseif sEliminar = "SI" then
      strsql = "Delete From Prodserv Where codprodserv = " & sID
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
      strsql = "Delete From atrmtemast Where atrmcodst = " & sID & " and atrmtipost='S'"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
    end if
  else
    dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
    dato = dato & "asp/EliminarSubtema.asp?ID=" & sID & "&Botones=1"
    response.redirect dato
  end if
end if

Volver = session("volver2")

set errores = cn.errors
if errores.count = 0 then
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>