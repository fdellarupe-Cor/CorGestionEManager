<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0


Tabla       = request.querystring("Tabla")

Volver      = request.querystring("Volver")
Volver      = ofv.convertircar(Volver,"_","&")
sUsuario    = session("Usuario")

on error resume next

set cn = ofv.conectar(ofv.strconn6)

wkspsep = ","

'strsql = "Select * From " & tabla
'set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  strsql = "Insert into " & tabla & " ("

valores = "values ("
delim=""
For Each x In Request.form
  strsql = strsql  &  delim & x 
  valores = valores  & delim  &"'" & Request.form(x)  &"'"
  delim=","
Next

strsql= strsql & ")"
valores=valores & ")"
strsqlx=strsql & valores
'response.write strsql  & valores
call ofv.crearconsultaEx(StrSqlx,cn,1,volver)

 

'call ofv.cerrarconsulta(rs)
call ofvcerrarconn(cn)
response.redirect volver 

%>