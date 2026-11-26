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
  he=request.querystring("he")
  Volver = request.querystring("Volver")
  Volver = Volver & "?Cod=" & sCod & "_CIRCUITO=" & SCIRCUITO & "_he=" & he
  sxopc1 = ""



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","N","N","S",CHR(9),CHR(9),"GRP-","USR-")

     sAccion = "WKFREASIGNADAV.ASP?OPCION=SETUSR&he=" & he &"&WKFCODH=" & sCod

     sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")

  '' response.write " OPCION=SETUSR " & chr(13) & "<br>"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botones Aceptar y Cancelar
     sAccion = "WKFREASIGNADAV.ASP?OPCION=NO&he=" & he & "&WKFCODH="& sCod
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  '' response.write " OPCION=NO " & chr(13) & "<br>"

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>