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

  directorio = Request.querystring("direct")
  nomdir = Request.querystring("nomdir")

  sHTML = ofv.FormHeader(session("FormN") & " - " & nomdir)
  'set application("nn")=ofv
  'response.write directorio
  'directorio=replace(directorio,"/","\")
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  response.write shtml

  sPath = Request.ServerVariables("PATH_INFO")
  spath2 = left(sPath,instrRev(sPath,"/",-1))
  spath2 = left(sPath2,len(sPath2)-1)
  spath2 = left(sPath2,instrRev(sPath2,"/",-1))




  'session("directorio")= "\" & directorio
  session("directorio")=  session ("DirUpload") & directorio
  
  'response.write session("directorio")

  response.write "<center><hr><h2>Seleccione los Archivos a transferir</h2><hr><br></center>"

  response.write "<form enctype=" & chr(34) & "multipart/form-data" & chr(34) & " action=NFiles_Save.asp method=post>"

    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 1:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "FILE1" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 2:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "FILE2" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 3:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "FILE3" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","","","","","","","<input type=" & chr(34) & "submit" & chr(34) & " value=" & chr(34) & "Transferir" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")





    sHTML =  ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  'response.write directorio
  response.write shtml
  response.write "</form>"
  response.write "</body>"

  response.write "</html>"

 '     Response.Write(  "Cliente=" & session("Cliente"))& "<br>"
 '     Response.Write(  "USRES=" & session("USRES"))& "<br>"
 '     Response.Write(  "Clave=" & session("Clave"))& "<br>"
 '     Response.Write(  "SegUsuario=" & session("SegUsuario"))& "<br>"
 '     Response.Write(  "SegNivel=" & session("SegNivel"))& "<br>"
 '     Response.Write(  "SegGrp=" & session("SegGrp"))& "<br>"
 '     Response.Write(  "SegPerfil=" & session("SegPerfil"))& "<br>"
 '     Response.Write(  "SkrOpc=" & session("SkrOpc"))& "<br>"
 '     Response.Write(  "CMRES=" & session("CMRES"))& "<br>"
 ''     Response.Write(  "USACC=" & session("USACC"))& "<br>"
 '     Response.Write(  "CONSET=" & session("CONSET"))& "<br>"


 'end if
end if
%>
