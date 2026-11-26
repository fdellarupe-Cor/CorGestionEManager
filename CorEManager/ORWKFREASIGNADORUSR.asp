<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sCod = request.querystring("Cod")
  Volver = request.querystring("Volver")
  Volver = Volver & "?Cod=" & sCod & "_CIRCUITO=" & SCIRCUITO
  sxopc1 = ""



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","N","N","S",CHR(9),CHR(9),"GRP-","USR-")



     sAccion = "ORWKFREASIGNADOR.ASP?OPCION=SETUSR"

     sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botones Aceptar y Cancelar
     sAccion = "ORWKFREASIGNADOR.ASP?OPCION=NO"
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")



  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>