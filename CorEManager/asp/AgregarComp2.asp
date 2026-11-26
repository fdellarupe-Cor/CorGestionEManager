<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nCodigo = request.querystring("Codcomp")
sTipo = request.querystring("Tipocomp")
volver = request.querystring("volver")

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "Select * From Compcompuesto"
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert Into Compcompuesto (compmaster,tipomaster"

For Each x In Request.form
' if x = "asoc" then
'    strsql = strsql & ",compasoc,tipoasoc"
'  else
    strsql = strsql & "," & x
'  end if
Next
strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "'"

For Each x In Request.form
  vNuevo = UCase(Request.Form(x))
 ' if x = "asoc" then
 '   s1 = ofv.queesx(rs("compasoc").type)
 '   if rs("compasoc").type = 200 then s1 = "'"
 '   s2 = ofv.queesx(rs("tipoasoc").type)
 '   if rs("tipoasoc").type = 200 then s2 = "'"
 '   i = instr(vNuevo,"_")
 '   codasoc = left(vNuevo,i-1)
 '   tipoasoc = mid(vNuevo,i+1)
 '   TipoValido = ofv.ValidarTipo(rs("compasoc"),codasoc)
 '   if TipoValido = False then exit for
 '   ValorNulo = ofv.EsNulo(rs("compasoc"),codasoc)
 '   if ValorNulo = True then exit for
 '   TipoValido = ofv.ValidarTipo(rs("tipoasoc"),tipoasoc)
 '   if TipoValido = False then exit for
 '   ValorNulo = ofv.EsNulo(rs("tipoasoc"),tipoasoc)
 '   if ValorNulo = True then exit for
 '   strsql = strsql & "," & s1 & codasoc & s1
 '   strsql = strsql & "," & s2 & tipoasoc & s2
 ' else
    s = ofv.queesx(rs(x).type)
   if rs(x).type = 200 then s = "'"
    TipoValido = ofv.ValidarTipo(rs(x),vNuevo)
    if TipoValido = False then exit for
    ValorNulo = ofv.EsNulo(rs(x),vNuevo)
    if ValorNulo = True then exit for
    strsql = strsql & "," & s & vNuevo & s
 ' end if
Next

volver = volver & "?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
if TipoValido = False or ValorNulo = True then
  on error resume next
  response.write ofv.ErrorTipo(volver,Request.Form(x),rs(nx).name)
  if err.number <> 0 then response.write ofv.ErrorTipo(volver,"","")
else
  strsql = strsql & ")"
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