<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

nID = request.querystring("ID")
t = request.querystring("t")
Nombre = request.querystring("Nombre")
sOpcion = request.querystring("Opcion")
sGrupoUsuario = request.querystring("grupousuario")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")

'response.write "gu " & sGrupoUsuario & " ID " & nid

xdefault = request.form("default")
xtipo = "E"

dim aatrm(9)
for xind=0 to 9
    aatrm(xind) = 1
next
xind = 0
xatrm = nid
xtab = "S"

set cn = ofv.conectar(ofv.strconn0)



  stabla = "ATRMEMP"



strsql = "SELECT * From ATRMTEMAST Where ATRMTIPOST = '" & T & "' AND ATRMCODST = " & sGrupoUsuario
strsql = strsql & " And ATRMCODES = " & NID
set rsa = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if rsa.eof then xtab = "N"
call ofv.cerrarconsulta(rsa)

xtipo = "A"





       aatrm(0) = request.form("ACCEDE")
       aatrm(1) = request.form("ACCEDED")


' se saco xq hace insert por defecto en 0
       'if aatrm(0) = aatrm(1)  then
       if false then
          if xtab <> "N" then
             strsql = "delete From ATRMTEMAST Where ATRMTIPOST = '" & T & "' AND ATRMCODST = " & sGrupoUsuario
             strsql = strsql & " And ATRMCODES = " & NID
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if
       else
          if xtab = "N" then
             strsql = "insert into ATRMTEMAST (ATRMTIPOST,ATRMCODST,ATRMCODES,ATRMVERST"
             strsql = strsql  & ") values("
             strsql = strsql  & "'" & T & "','" & sGrupoUsuario & "','" & NID & "'"
             strsql = strsql  & "," & aatrm(0) & ")"
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          else
             strsql = "update ATRMTEMAST set "
             strsql = strsql & " ATRMVERST = " & aatrm(0)
             strsql = strsql & " Where ATRMTIPOST = '" & T & "' AND ATRMCODST = " & sGrupoUsuario
             strsql = strsql & " And ATRMCODES = " & NID
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if
       end if


  ' RESPONSE.WRITE STRSQL



if cn.errors.count = 0 then
  volver = volver & "&Opcion=" & sGrupoUsuario & "&t=" & t & "&Nombre=" & Nombre
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconn(cn)

%>