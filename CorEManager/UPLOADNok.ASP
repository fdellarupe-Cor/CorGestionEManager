<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<% Response.Buffer = TRUE %>
<% if (Len(Request.ServerVariables("LOGON_USER")) = 0 ) then %>
	<% Response.Status = "401 Unauthorized" %>
	<HTML><BODY><B>Error: Access is denied.</B><P></BODY></HTML>
<% else

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

  Xruta  = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT") & left(sPath,instrRev(sPath,"/",-1))
'''  RESPONSE.WRITE XRUTA & "<BR>"

  Xruta2 = Request.ServerVariables("SERVER_NAME") & sPath2 & "coressdll/"

'  response.write "path2:" & Xruta2 & "<br>"

  response.write "<center><hr><h2>Seleccione los Archivos a transferir</h2><hr><br></center>"

  'response.write "<form target=_new enctype=" & chr(34) & "multipart/form-data" & chr(34) & " action=" & chr(34) & "../CORESSDLL/cpshost.dll?PUBLISH?HTTP://" &  Xruta & "repost.asp" & chr(34) & " method=post>"
   response.write "<form target=_new enctype=" & chr(34) & "multipart/form-data" & chr(34) & " action=" & chr(34) & "../CORESSDLL/cpshost.dll?PUBLISH "&chr(34)&" method=post>"
   
    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 1:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "my_file" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 2:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "my_file" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 3:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "my_file" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt","","","","","","Destino (URL):","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","http://" & directorio  & " <input class=nt name=" & chr(34) & "TargetURL" & chr(34) & " value=" & chr(34) & "http://" & directorio  & chr(34) & " size=100  type=" & chr(34) & "hidden" & chr(34) & " >","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","","","","","","","<input type=" & chr(34) & "submit" & chr(34) & " value=" & chr(34) & "Transferir" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    sHTML =  ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")


  response.write shtml
  response.write "</form>"
  response.write "</body>"

  response.write "</html>"
 end if 
end if 
%>
