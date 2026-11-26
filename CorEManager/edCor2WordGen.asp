<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
'  response.buffer = true
  server.scripttimeout = 300
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Id")
  sTipo = request.querystring("Tipobj")
  sTitulo = request.querystring("titulo")
  sperfil = request.querystring("perfil")


  sTitulo = replace(stitulo,"_"," ")
  shtml=""

'  response.write ofv.FormHeader("Conversion a Word")



'  response.write "<br><br><p align=center ><font face=tahoma size=2 color=#993333><b>Conversion de " & stitulo & ", espere un momento</b></font></p>"

'  response.flush

  if len(stitulo) > 0 then 

  '''set oSRV = server.createobject("CorWebQueueXPX.WordServer")

  set oPrtG = server.createobject("CorWebQueueXPX.Admin")

  Set oPrtG.Request = Request
  Set oPrtG.Server = Server
  Set oPrtG.Response = Response
  Set oPrtG.Session = Session
  Set oPrtG.Ofv = Ofv


         sPath = Server.mappath(Request.ServerVariables("SCRIPT_NAME"))
         sPath = Mid(sPath, 1, InStrRev(sPath, "\", -1))
         sPath = Mid(sPath, 1, InStrRev(sPath, "\", -1))
   '      sPathx = replace(sPath,"\","/")
         sPathx = sPath
         sPath = left(sPath, InStrRev(sPath, "\", len(sPath)-1))
         sOutput = sPath & "CorEss20\CorDownload\FW" & sTipo & nCodigo & session.sessionid & ".doc" 

         sOutputx = sOutput


 ' sOutputx = replace(sOutput,"\","/")


  'response.write "<br><br>" & sOutputx
  'response.write "<br><br>" & sPathx
  
  ' response.write oPrtG.Generar(clng(nCodigo), cstr(sTipo), "Esquemas Normativos", cstr(sperfil),cstr(sOutputx),cstr(sPathx))


  sMsg = oPrtG.WordQueue(clng(nCodigo), cstr(sTipo), "Esquemas Normativos", cstr(sperfil),cstr(sOutputx),cstr(sPathx),Request.ServerVariables("REMOTE_ADDR"),session("usuario"))

  response.redirect "edcor2wordSrch.asp?xitem=" & smsg

  sHTML = "<html> "
  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "WDCV" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"
  sHTML = sHTML  & "<body onload=AbrirCons('edcor2wordSrch.asp?xitem=" & smsg & "'); >"

  sHTML = sHTML  & "<p>" & oPrtG.ofv.ADbUidl & "</p>"
  sHTML = sHTML  & "<p>" & oPrtG.ofv.ADbPwdl & "</p>"
  sHTML = sHTML  & "<p>" & oPrtG.session.sessionid & "</p>"

  sHTML = sHTML  & "</body></html>"


 ''  response.write sHTML

  end if







end if

'response.write sHTML
'response.flush

%>