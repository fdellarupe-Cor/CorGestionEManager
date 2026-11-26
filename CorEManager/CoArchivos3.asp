<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
   xext = " "
  nArchivo = request.querystring("Archivo")
  xinstr = instr(nArchivo,".")
  if xinstr > 0 then
     xext = mid(nArchivo,xinstr + 1,len(nArchivo) - xinstr)
  end if
  set fff = server.createobject("scripting.filesystemobject")
  Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
  XferTarget1 = mid(Xfer,1,instrrev(Xfer,"\",-1))
 '' XferTarget1 = XferTarget1 & "Archivos\" & nArchivo
  XferTarget1 = XferTarget1 & nArchivo

  sPath = Request.ServerVariables("PATH_INFO")

   if  SESSION("URLPORT")  = "N" then
    Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1))
   else
    Xruta = Request.ServerVariables("SERVER_NAME")& ":" & Request.ServerVariables("SERVER_PORT") & left(sPath,instrRev(sPath,"/",-1))
   end if

''  Xruta = Xruta &  "Archivos/" & nArchivo

 Xruta = Xruta &  nArchivo


 ' response.write xruta
  'response.write xruta
  if nArchivo <> "" then

 ''  nArchivo = "Archivos/" & nArchivo


   'response.write XferTarget1
  ' Set fff1 = fff.openTextFile(trim(XferTarget1), 1, -1)
   'if fff1.AtEndOfstream <> true then
      select case  Ucase(xext)
             case  "JPG","JPEG","HTM","HTML","BMP","GIF","TIFF","XLS","ASP","PDF","PPT","DOC","MPG","MPEG"

               '    response.write "<html><body bgcolor=#cccccc onload='document.all.VD.click();' >"
                   response.write ofv.MenuHeader(""," onload=""document.getElementById('VD').click();"" ")
                   response.write "<br><br><center><a class=nt name=VD id=VD href=" & chr(34) & TipoSite("") & Xruta & chr(34)
                   response.write " ><font size=2 face=tahoma ><b>Cargando documento, espere un momento...</font></a><br><br></center></body></html>"

              '     response.redirect chr(34) & nArchivo & chr(34)
              '     response.redirect nArchivo
             case else
                   response.write "<html><body>"
                 '  response.write " </body></html>"
                   response.write CorStyler2()
                   response.write "<br><br><center><a class=nt href=" & chr(34) & TipoSite("")  & Xruta & chr(34)
                   response.write " target=_blank><font size=2 face=tahoma ><b>Seleccione aqui para visualizar</font></a><br><br></center></body></html>"

      end select

 ' else
 '  response.write ofv.MenuHeader("","  ")
 '  response.write "<center><font size=2 face=tahoma color=#993333>"
 '  response.write "<br><br>Documento Sin Informacion</font></center></body></html>"
 ' end if
  else
   response.write ofv.MenuHeader("","  ")
   response.write "<center><font size=2 face=tahoma color=#993333>"
   response.write "<br><br>Documento Sin Informacion</font></center></body></html>"
  end if

end if



%>