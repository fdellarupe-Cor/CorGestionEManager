<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  salng = request.querystring("A")
  sbord = request.querystring("B")
  sreduc = request.querystring("R")
  simg = request.querystring("S")

  response.write ofv.menuheader("#ffffff"," onload=" & chr(34) & "window.title='Edicion de Contenido - Vista Preliminar de Imagen';" & chr(34))


  response.write ofv.GenTabla("","","","100% border=0 ","","0","","0","0","#ffffff","","","no")
  Celda =  ofv.GenCelda("","tt","center","","","","","VISTA PRELIMINAR DE IMAGEN: " & SIMG,"si")
  response.write ofv.GenRow("","","","","","","",celda,"si")
  Celda =  ofv.GenCelda("","cc","","","","","","<br>&nbsp;<b>Propiedades</b><br>","si")
  response.write ofv.GenRow("","","","","","","",celda,"si")

 ' IMAGENES            --------------------------------------------------

     xred = "100%"
     xbord = "NO"
     xalng = "IZQUIERDA"

      if sreduc <> "" then
         xred = sreduc & "%"
         sreduc = " width=" & sreduc & "% "
      end if

      if sbord <> "" then
         if sbord = "SI" then
            sbord = " border=1 "
            xbord = "SI"
         end if
      end if

      if salng <> "" then
         xalng = salng
         if salng = "CENTRADO" then salng = " center"
         if salng = "DERECHA" then salng = " right"
         if salng = "IZQUIERDA" then salng = " left"
      end if

  Celda =  ofv.GenCelda("","ut","","","","","","&nbsp;Mostrar en: " & xred,"si")
  response.write ofv.GenRow("","","","","","","",celda,"si")
  Celda =  ofv.GenCelda("","ut","","","","","","&nbsp;Marco: " &  xbord,"si")
  response.write ofv.GenRow("","","","","","","",celda,"si")
  Celda =  ofv.GenCelda("","ut","","","","","","&nbsp;Alineaci&oacute;n: " & xalng,"si")
  response.write ofv.GenRow("","","","","","","",celda,"si")
  Celda =  ofv.GenCelda("","cc","","","","","","<br>&nbsp;<b>Imagen</b><br>","si")
  response.write ofv.GenRow("","","","","","","",celda,"si")

if session("EMPMUN") = "S" then
         slink = "<img src=" & chr(34) & "Emp" & SESSION("D" & session("ncliente") & "-0" ) & "/estructura/" & simg  & chr(34) & " "
         slink = slink & sreduc & sbord & " >"
else
		 slink = "<img src=" & chr(34) & "estructura/" & simg  & chr(34) & " "
         slink = slink & sreduc & sbord & " >"
end if


      Celda =  ofv.GenCelda("","",salng,"","","","",slink,"si")

      response.write ofv.GenRow("","","","","","","",celda,"si")
      response.write ofv.GenFTabla() & "</body></html>"

  sHTML = ""
end if

response.write sHTML

%>