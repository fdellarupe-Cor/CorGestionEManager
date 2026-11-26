<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  response.write sHTML
else
  response.expires=0

  XarchivoV = request.querystring("archivo")
  TipObjeto = request.querystring("TipObjeto")
  CodObjeto = request.querystring("CodObjeto")
  opc = request.querystring("opc")

  Xarchivo = request.form("nombre")
  Xarchivo = ofv.convertircar(Xarchivo,"_","\")
  Xdatos = request.form("datos")
  XTxtVacio=Trim(request.form("Txtvacio"))

  set fso = server.createobject("scripting.filesystemobject")

  if len(XTxtVacio)=0 or XTxtVacio="" then
     if fso.FileExists(Xarchivo) then
        fso.deleteFile(Xarchivo)
     end if
  else
    Set fa = fso.openTextFile(Xarchivo, 2, -1)
    fa.write Xdatos
    fa.close
  end if

'  if ofv.explorador = "MSIE" then
    wlink = "edtransfer.asp?Codobjeto=" & CodObjeto & "&Tipobjeto=" & Tipobjeto & "&opc=" & opc & "&archivo="
 ' else
 '   wlink = "edtransferNSC.asp?Codobjeto=" & CodObjeto & "&Tipobjeto=" & Tipobjeto & "&opc=" & opc & "&archivo="
 ' end if

  response.redirect wlink &  trim(XarchivoV)
end if

%>