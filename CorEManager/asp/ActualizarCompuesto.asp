<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nCodigo = request.querystring("Codcomp")
sTipo = request.querystring("Tipocomp")
ID = request.querystring("ID")
ID = ofv.convertircar(ID,"*"," ")  
Tabla = request.querystring("Tabla")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")

on error resume next

set cn = ofv.conectar(ofv.strconn2)

strsql = "select * from " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
s = ofv.QueEsx(rs(0).type)
if rs(0).type = 200 then s = "'"
strsql = "select * from " & tabla & " where " & rs(0).name & "=" & s & id & s
call ofv.cerrarconsulta(rs)
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
strsql= "update " & tabla & " set "

bEjecutar = False
For Each yy In Request.form
x = lcase(yy)
  vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  if x = "codcomp" then x = "compasoc"
  TipoValido = ofv.ValidarTipo(rs(x).name,vNuevoValor)
  if TipoValido = False then exit for
  ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
  if ValorNulo = True then exit for
  s = ofv.QueEsx(rs(x).type)
  if rs(x).type = 200 then s = "'"
  if vNuevoValor <> "NSNC" then
    bEjecutar = True
    strsql=  strsql & x & " = " & s & vNuevoValor & s & " ,"
  else
    bEjecutar = True
    if s = "'" then
      strsql=  strsql & x & " = ' ',"
    else
      strsql=  strsql & x & " = 0,"
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

  response.redirect "errortipo.asp?volver=" & volver & "&Columna=" & rs(x).name & "&Dato=" & vNuevoValor
else
  s = ofv.QueEsx(rs(0).type)
  if rs(0).type = 200 then s = "'"
  strsql = mid(strsql,1,len(strsql)-1) & " where " & rs(0).name & "=" & s & id & s
  if bEjecutar then 
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
  end if 

  if cn.errors.count = 0 then
    response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
  end if 
end if 

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>