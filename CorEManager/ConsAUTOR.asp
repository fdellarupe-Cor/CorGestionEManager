<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0
nCodigo = request.querystring("Codigo")
sTipo = request.querystring("Tipo")
sPerfil = request.querystring("Perfil")
sVolver = request.querystring("Volver")
sDatos = ""
if request.querystring("Datos") <> "" then
  sDatos = ofv.convertircar(request.querystring("Datos"),"*","&")
end if

Solapas = 8

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then   sEsquema = rs(1)

  call ofv.cerrarconsulta(rs)

  Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  Solapas = 1
  for z = 2 to 12
    if rs(z) Then
       Solapas = Solapas + 1
    end if
  next
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

sHTML = "<html><script language=JavaScript for=window event=onload >" & chr(13)
sHTML = sHTML & "window.moveTo(0,0);" & chr(13)
If BrowserName() = "IE" OR BrowserName() = "CHROME" then
sHTML = sHTML & "window.document.title =" & chr(34) & "COR EManager - Consulta de Esquemas Normativos" & chr(34) & ";" & chr(13)
end if
sHTML = sHTML & "</script>" & chr(13)

if Solapas > 7 then
   sHTML = sHTML & "<frameset rows=65,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
else
   sHTML = sHTML & "<frameset rows=45,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
end if

sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "BarraEN.asp?Tipo="
sHTML = sHTML & sTipo & "&Perfil=" & sPerfil & "&Codigo=" & nCodigo & "&Funcion=CA" & sDatos
sHTML = sHTML & "&Volver=" & sVolver & chr(34) & " NAME=Barra SCROLLING=auto>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Principal "
sHTML = sHTML & "src=" & chr(34) & "CoGeneral.asp?Tipobjeto=" & sTipo
sHTML = sHTML & "&Codobjeto=" & nCodigo & chr(34) & " SCROLLING=auto></frameset>"
sHTML = sHTML & "<body><head><title>COR EManager - Edicion de Esquemas Normativos</title></head></body></html>"

response.write sHTML

%>