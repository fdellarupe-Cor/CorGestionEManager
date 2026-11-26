<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  shtml = ofv.formheader(session("FormN"))


  direct  = request.querystring("direct")
  nomdir  = request.querystring("nomdir")
  rutdir  = request.querystring("rutdir")
  Browdir = request.querystring("Browdir")




'  response.write direct & " 0<br>"
  Set fso = CreateObject("Scripting.FileSystemObject")

  sPath = Request.ServerVariables("PATH_INFO")
'  response.write spath & " 1<br>"
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1) - 1)
' response.write Xruta & " 2<br>"
 sPath = server.mappath(sPath)
'  response.write spath & " 3<br>"
    sPath = replace(sPath,"COREManager\cordscrBrow.asp","",1,-1,1) & "CORESS20\ASP"
  response.write spath & " 4<br>"
  Set Folder = fso.getfolder(sPath)
  Set Files = Folder.Files
 'response.write Browdir
 ' sfiles=split(getfilesinfolder(Browdir,""),"|")
  
 ' response.write 
  sHTML = sHTML & ofv.GenTabla("","","","50%","","0","0","0","0","","","","no")
  sHTML = sHTML & "<font color=#DCDCDC face=tahoma size=1><b>"
 h1 = ofv.GenCelda("","ttt colspan=2","","","","","",Browdir,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
 h1 = ofv.GenCelda("","ttr","","","","","","descripcion","si")
 h2 = ofv.GenCelda("","ttr","","","","","","creacion","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1 & H2,"si")
  sHTML = sHTML & "</b></font>"

  xTrColor = "#F9F9F9" 

  For Each F in Files
 ' response.write Browdir
 '  for f=0 to ubound(sfiles)-1

     xVersion = ""
   '  response.write sfiles(f) & "<br>"
   '  file=split(sfiles(f),"###")
  '   saccion = "<a class=nt href=" & chr(34) & browdir & "/" & file(0)  & chr(34)
  '   saccion = saccion & "  target=_blank ><font color=#993333 face=corflaticon size=4 >3</font>&nbsp;&nbsp;" & file(0) & "</a>"
     h1      = ofv.GenCelda("","vt","","","","","",ucase("- " & f.name & ""),"si")
     h2      = ofv.GenCelda("","ut","","","","","","-","si")
     sHTML   = sHTML & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","22","","",h1 & H2,"si")

     if  xTrColor = "#FFFFFF" then
         xTrColor = "#F9F9F9" 
     ELSE
         xTrColor = "#FFFFFF" 
     END IF

  Next



    sHTML = sHTML & ofv.GenFTabla()


  Agregar = ""
  Buscar  = ""

  sAccion = "Cordscr2.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir & "&OPC=NO"
  Celda   = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>