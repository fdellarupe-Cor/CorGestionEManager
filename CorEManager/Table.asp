<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader(session("FormN"))

  posI = 0
  sForm = "form" & posi

  sInput = "Norma"
  Celda = ofv.gencelda("","TTT COLSPAN=3 ","","","","","",sInput,"si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.gencelda("","tt","","20%","","","","Perfil","si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.gencelda("","tt","","","","","","Nivel","si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.GenCelda(""," ttt colspan=5","","","","","","  CIRCUITOS ","si")
  Celda = Celda & ofv.GenCelda("","ttt colspan=5","","","","",""," CIRCUITO DEFAULT ","si")
  Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","2","","",Fila,"si") & "<br>"

  sHTML = sHTML & table
  sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

%>