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
  sacc = Request.querystring("acc")
  idpedido = Request.querystring("id")
  seg=REQUEST.querystring("seg")
  sVuel = REQUEST.querystring("vuel")

  sHTML = ofv.FormHeader(session("FormN") & " - " & nomdir)
  'set application("nn")=ofv
  'response.write directorio
  'directorio=replace(directorio,"/","\")
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form
  response.write shtml

  sPath = Request.ServerVariables("PATH_INFO")
  spath2 = left(sPath,instrRev(sPath,"/",-1))
  spath2 = left(sPath2,len(sPath2)-1)
  spath2 = left(sPath2,instrRev(sPath2,"/",-1))

'reeponse.write spath2



  session("directorio")= spath2
  session("directorio")=  session ("DirUpload") & directorio

  'response.write session("directorio")

  response.write "<center><hr><h2>Seleccione los Archivos a transferir</h2><hr><br></center>"
  if  SESSION("FILETIPO")  = "L" then
  response.write "<form name=fx id=fx enctype=" & chr(34) & "multipart/form-data" & chr(34) & " action=NFiles_Save.asp?seg=" & seg &"&acc=" & sacc & "&id=" & idpedido & " method=post>"
  else
  response.write "<form name=fx id=fx  enctype=" & chr(34) & "multipart/form-data" & chr(34) & " action=fileserver/asp/NFiles_Save.asp method=post>"
  end if

    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 1:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "FILE1" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 2:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "FILE2" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
    if  SESSION("FILETIPO")  = "R" then
    SINPUT="<INPUT TYPE=HIDDEN VALUE='"& directorio&"/' NAME=DIRC>"
    else
    SINPUT=""
    end if
    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 3:","si")
    Celda = Celda & ofv.gencelda("","cc","","","","","","<input class=nt name=" & chr(34) & "FILE3" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
    SINPUT2=""
    IF idpedido <> "" THEN
     SINPUT2 = "<INPUT TYPE='BUTTON' VALUE='    Volver    ' CLASS='BT' onClick=document.getElementById('fped').submit();>"
     Celda = ofv.gencelda("","XX ALIGN=RIGHT","","","","","","","si")
    ELSE


    Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
    END IF

    Celda = ofv.gencelda("","xx colspan=2 align=center","","","","","",SINPUT2 & "<input class=bt type=" & chr(34) & "button" & chr(34) & " onClick=document.getElementById('fx').submit(); value=" & chr(34) & "Transferir" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




    sHTML =  ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  'response.write directorio
  SHTML =shtml & SINPUT

  SHTML =shtml & "</form>"
  if sVuel = "PED" then
   		sAccion="CREARPEDIDO.asp?idpedido=" & idpedido & "&seg=" & seg & "&PEDIDO=" & SACC
  else
  		sAccion="CrearNormativa.asp?idpedido=" & idpedido & "&seg=" & seg & "&PEDIDO=" & SACC
  end if
   Formulario = ofv.GenForm("fped id=fped","",sAccion,"","POST","","","si")
  Agregar = ""
    Buscar  = ""

  '  sAccion = "PerfilesdeEsquemaNormativo.asp?CantRegMover=" & CantRegMoverx
  '  Celda   = ofv.BotonVolver(sAccion)
  ' sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",CELDA)
  ' sHTML = sHTML & "<script>document.getElementById('MTH').style.display='none';</script></body></html>"
  RESPONSE.WRITE SHTML & Formulario
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
