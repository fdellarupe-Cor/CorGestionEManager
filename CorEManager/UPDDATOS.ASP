<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Set cn = ofv.conectar(ofv.strconn0)

  datos=request.form("inforeg")
  adatos=split(datos,"|")
'  pos=1
'  for i=0 to ubound(adatos)
  ' xcols=split(adatos(i),"|")
   sql="update novedtitex set posy=" & adatos(1) & ",posx=" & adatos(2) &",height=" &adatos(3)& ",width=" & adatos(4) & " where objid= " & adatos(0)
   response.write sql & "<br>"
   call ofv.crearconsultaEx(sql, cn, 1, parametros)
  ' pos=pos + 1
 ' next
  set cn = ofv.conectar(ofv.strconn2)



end if

Response.Write sHTML

%>