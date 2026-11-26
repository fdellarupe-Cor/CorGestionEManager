<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo = request.querystring("Tipo")
sNombre = request.querystring("Nombre")

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "select * from Complementos"
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert into Complementos (Tipocomp"

For Each x In Request.form
  strsql = strsql & "," & x
Next
strsql = strsql & ") Values ('" & sTipo & "'"

For Each x In Request.form
  TipoValido = ofv.ValidarTipo(rs(x),Request.Form(x))
  if TipoValido = False then exit for
  ValorNulo = ofv.EsNulo(rs(x),Request.Form(x))
  if ValorNulo = True then exit for
  if rs(x).name = "datoformul" and Request.Form(x) = "nsnc" then
    strsql = strsql & ",0"
  else
    s = ofv.QueEsx(rs(x).type)
    if rs(x).type = 200 then s = "'"
 '''strsql = strsql & "," & s & UCase(Request.Form(x)) & s
    strsql = strsql & "," & s & Request.Form(x) & s
  end if
Next

volver = "../CompTree.asp"

'RESPONSE.WRITE VOLVER

if TipoValido = False or ValorNulo = True then
  response.redirect "errortipo.asp?volver=" & volver & "&Columna=" & rs(x).name & "&Dato=" & Request.Form(x)
else
  strsql = strsql & ")"
  'RESPONSE.WRITE STRSQL
  call ofv.crearconsultaEx(StrSql,cn,1,volver)

  if cn.errors.count = 0 then
    response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
  end if
end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>