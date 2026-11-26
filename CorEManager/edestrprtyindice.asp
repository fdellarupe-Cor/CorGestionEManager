<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  volver = request.querystring("volver")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ID = request.querystring("ID")

  sHTML = ofv.Menuheader(Session("CMFMENU"),"")

  sHref = "edestrprtycont.asp?volver=" & volver & "&Codobjeto=" & nCodigo & "&Tipobjeto="
  sHref = sHref & sTipo & "&ID=" & ID
  sHref1 = sHref & "&Tipo=Titulo"
  sHref2 = sHref & "&Tipo=Subtitulo1"
  sHref3 = sHref & "&Tipo=SubTitulo2"
  sHref4 = sHref & "&Tipo=SubTitulo3"
  sHref5 = sHref & "&Tipo=SubTitulo4"
  sHref6 = sHref & "&Tipo=SubTitulo5"
  sHref7 = sHref & "&Tipo=Texto"
  Celda = ofv.GenCelda("","tt","center","1%","","","","Propiedades Generales","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  Celda = ofv.GenCelda("","","center","1%","1","","","","si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sLink = ofv.GenLink("","nt",sHref1,"Contenido","","",Ucase("Titulo"),"si")
  Celda = ofv.GenCelda("","cc","","1%","","","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sLink = ofv.GenLink("","nt",sHref2,"Contenido","","",Ucase("Subtitulo 1"),"si")
  Celda = ofv.GenCelda("","cc","","1%","","","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sLink = ofv.GenLink("","nt",sHref3,"Contenido","","",Ucase("Subtitulo 2"),"si")
  Celda = ofv.GenCelda("","cc","","1%","","","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sLink = ofv.GenLink("","nt",sHref4,"Contenido","","",Ucase("Subtitulo 3"),"si")
  Celda = ofv.GenCelda("","cc","","1%","","","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sLink = ofv.GenLink("","nt",sHref5,"Contenido","","",Ucase("Subtitulo 4"),"si")
  Celda = ofv.GenCelda("","cc","","1%","","","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sLink = ofv.GenLink("","nt",sHref6,"Contenido","","",Ucase("Subtitulo 5"),"si")
  Celda = ofv.GenCelda("","cc","","1%","","","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sLink = ofv.GenLink("","nt",sHref7,"Contenido","","",Ucase("Texto"),"si")
  Celda = ofv.GenCelda("","cc","","1%","","","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","10","#993333","",Fila,"si")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>