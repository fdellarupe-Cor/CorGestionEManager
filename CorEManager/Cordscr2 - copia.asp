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



  direct = request.querystring("direct")
  nomdir = request.querystring("nomdir")
  rutdir = request.querystring("rutdir")
  opc = request.querystring("OPC")
'response.write request.querystring

  if opc = "" then opc = "NO"
  if newdir = "" then newdir = "NO"

  Xruta = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
  sxx   = Request.ServerVariables("URL")

  sxx   = left(sxx,instrRev(sxx,"/",-1) - 1)

  xurld = replace("\\" & Xruta & sxx,"/","\") & "\" & direct
  'response.write nomdir & " <br>"

 ' xurlf = server.mappath(xurld)
 ' response.write xurlf & " <br>"

'  response.write direct & " 0<br>"
 ' Set fso = server.CreateObject("Scripting.FileSystemObject")
  'response.write rutdir
  sPath = Request.ServerVariables("PATH_INFO")
'  response.write spath & " 1<br>"
  Xruta = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")  & left(sPath,instrRev(sPath,"/",-1) - 1)
' response.write Xruta & " 2<br>"
 sPath = server.mappath(sPath)
'  response.write spath & " 3<br>"
    sPath = replace(sPath,"cordscr2.asp","",1,-1,1) & direct
'  response.write spath & " 4<br>"
'response.write spath
  'Set Folder = fso.getfolder("\\cor03\pci")
'response.write spath
 ' Set Folder = fso.getfolder(xurld)
 ' set sfolder = Folder.SubFolders
  'Set Files = Folder.Files

  sHTML = sHTML & ofv.GenTabla("","","","650","","0","0","0","0","","","","no")
  sHTML = sHTML & "<font color=#DCDCDC face=tahoma size=1><b>"
 h1 = ofv.GenCelda("","ttt colspan=2","","","","","",nomdir,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
 h1 = ofv.GenCelda("","ttr","","","","","","descripcion","si")
 h2 = ofv.GenCelda("","ttr","","","","","","tipo","si")
    sHTML = sHTML & ofv.GenRow("","","","","22","","",h1 & H2,"si")
  sHTML = sHTML & "</b></font>"

  saccion = "<a class=nt href=uploadn.asp?direct=" & rutdir & "&nomdir=" & nomdir & "  target=_blank ><font color=#993333 face=CorFlatIcon size=2>1</font>&nbsp;&nbsp;<b>" & Ucase(direct) & "</b></a>"
  sBrow = "<a class=nt href=CordscrBrow.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir & "&browdir=" & direct & "   ><b>+</b></a>"
  h1 = ofv.GenCelda("","ut","","","","","",sBrow & "&nbsp;" & saccion,"si")
  h2 = ofv.GenCelda("","ut","","","","","","Raiz","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1 & H2,"si")

  llena = space(20)
  llena = replace(llena," ","&nbsp;")

  if  opc = "SI" then
      newdir = request.form("newdir")
      if  newdir = "" then newdir = "NO"
      if  newdir <> "NO" then
          'sfolder.add(newdir)
         ' RESPONSE.WRITE direct
          IF SESSION("FILETIPO")  = "L" THEN
           call getdata( "","AF",direct,newdir)
          ELSE
           call ADDFOLDER(direct,newdir,"")
          END IF
      end if
  end if

  xTrColor = "#F9F9F9"
  'response.write direct
  sfolder=split(getfolders(rutdir,""),"|")
 ' response.write (getfolders(direct,"")
 ' For Each F in sfolder
 '' rutdir="server-41:80/bb"
   for f=0 to ubound(sfolder)-1

     ' saccion = "<a class=nt href=uploadn.asp?direct=" & rutdir & "/" & sfolder(f) & "&nomdir=" & nomdir
    '  saccion = "<a class=nt href=../fileserver/asp/uploadn.asp?direct=" & direct & "/" & sfolder(f) & "&nomdir=" & nomdir
      saccion = "<a class=nt href=uploadn.asp?direct=" & direct & "/" & replace(sfolder(f),chr(32),"%20") & "&nomdir=" & nomdir
      saccion = saccion & "  target=_blank >" & llena & "<font color=#993333 face=CorFlatIcon size=2>1</font>&nbsp;&nbsp;/" & sfolder(f) & "</a>"
      sBrow   = "<a class=nt href=CordscrBrow.asp?direct=" & direct & "&nomdir=" & nomdir
      sbrow   = sBrow & "&rutdir=" & rutdir & "&browdir=" & direct & "/" & replace(sfolder(f),chr(32),"%20") & "   ><b>+</b></a>"
      h1      = ofv.GenCelda("","ut","","","","","",sBrow & "&nbsp;" & saccion,"si")
      h2      = ofv.GenCelda("","ut","","","","","","dir","si")
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

  sAccion = "imgfiles.asp"
  Celda   = ofv.BotonVolver(sAccion)

  xxxx    = CorGenBotonXp(3, "NO", "NO", "NO", "NO", "Agregar Subdirectorio", "NO", "Agregar Subdirectorio")
  sAccion = "Agregarsubdir.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir
  Agregar = ofv.BotonAgregar(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>