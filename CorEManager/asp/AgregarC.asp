<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nID = request.querystring("ID")
nCodigo = request.querystring("Codcomp")
sTipo = request.querystring("Tipocomp")
Tabla = request.querystring("Tabla")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "select * from " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert into " & tabla & " (Codcomp,Tipocomp"
if tabla = "Compdet" then strsql = strsql & ",codestr"

For Each x In Request.form
  strsql = strsql & "," & x
Next
strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "'"
if tabla = "Compdet" then strsql = strsql & "," & nID

For Each x In Request.form
''vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
  TipoValido = ofv.ValidarTipo(rs(x),vNuevoValor)
  if TipoValido = False then exit for
  ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
  if ValorNulo = True then exit for
  s = ofv.QueEsx(rs(x).type)
  if rs(x).type = 200 then s = "'"
  if vNuevoValor <> "NSNC" then
    strsql = strsql & "," & s & vNuevoValor & s
  else
    if s = "'" then
      strsql=  strsql & ",''"
    else
      strsql=  strsql & ",0"
    end if
  end if
Next

if volver <> "" then
  if right(volver,3) = "asp" then
    volver = volver & "?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
  else
    volver = volver & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
  end if
end if

if TipoValido = False or ValorNulo = True then
  response.redirect "errortipo.asp?volver=" & volver & "&Columna=" & rs(x).name & "&Dato=" & Request.Form(x)
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