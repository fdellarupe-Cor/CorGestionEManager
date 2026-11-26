<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
 Response.Buffer = TRUE 

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(session("FormN"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  response.write shtml

  response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br></center>"


    Celda = ofv.gencelda("","tt","","","","","","Destino:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","",Request.Form("TargetURL"),"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")


  For I = 1 To Request.Form("FileName").Count

    Celda = ofv.gencelda("","tt","","","","","","Archivo copiado:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<b>" & Request.Form("FileName")(I) & Request.Form("FileExtention")(I) & "</b>","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt","","","","","","Path:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<b>" & Request.Form("FilePath")(I) & "</b>","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt","","","","","","Tamaño:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<b>" & Request.Form("FileSize")(I) & "</b>","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


  Next


    sHTML =  ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  response.write shtml

  response.write "</BODY></HTML>"

 end if
%>