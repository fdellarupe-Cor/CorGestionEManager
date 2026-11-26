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

  ofv.ObtenerAtributos Session("Form"),""
    bAgregar = ofv.AAgregar
    bModificar = ofv.AModificar
    bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
 aplcheck()
  DIM  direct,nomdir,rutdir,Browdir
  direct  = request.querystring("direct")
  nomdir  = request.querystring("nomdir")
  rutdir  = request.querystring("rutdir")
  Browdir = request.querystring("Browdir")
  OP=REQUEST.querystring("OP")

  IF  request.FORM.COUNT >1 AND OP="DEL" THEN
    direct  = request.FORM("direct")
    nomdir  = request.FORM("nomdir")
    rutdir  = request.FORM("rutdir")
    Browdir = request.FORM("Browdir")
 '    response.write request.form("direct")
    ARCH=REQUEST.FORM("ARCH")

  END IF
'response.write request.querystring
   sImageELN    = "<font face='Corflaticon' size=2 color='#dddddd'>&#0127;</font>"
   sImageELS    = "<font face='CorFlatIcon' size=2 color='#bb0000'>&#4;</font>"
   wktitdelete = "Confirma la eliminacion???."

'  response.write direct & " 0<br>"
  'Set fso = CreateObject("Scripting.FileSystemObject")

  sPath = Request.ServerVariables("PATH_INFO")
'  response.write spath & " 1<br>"
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1) - 1)
' response.write Xruta & " 2<br>"
 sPath = server.mappath(sPath)
'  response.write spath & " 3<br>"
    sPath = replace(sPath,"cordscrBrow.asp","",1,-1,1) & Browdir

 if OP="DEL" THEN
  delfile(replace(Browdir,"/","\") & "\" & arch)
  'response.write replace(Browdir,"/","\") & "\" & arch
 END IF
  sfiles=split(getfilesinfolder(Browdir,""),"|")

 ' response.write
  sHTML = sHTML & ofv.GenTabla("","","","50% ","","0","0","0","0","","","","no")
  sHTML = sHTML & "<font color=#DCDCDC face=tahoma size=1><b>"
  h1 = ofv.GenCelda("","ttt colspan=2","","","","","",REPLACE(replace(Browdir,"%20","&nbsp;"),"\","/"),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
  h3 = ofv.GenCelda("","ttr","","","","","","*","si")
  h1 = ofv.GenCelda("","ttr","","","","","","descripcion","si")
  h2 = ofv.GenCelda("","ttr","","","","","","creacion","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h3 & h1 & H2,"si")
  sHTML = sHTML & "</b></font>"

  xTrColor = "#F9F9F9"

 ' For Each F in Files
 ' response.write Browdir


  sfolder=split(getfolders(Browdir,""),"|")
  ' response.write (getfolders(direct,"")
  ' For Each F in sfolder
  '' rutdir="server-41:80/bb"
    for f=0 to ubound(sfolder)-1
          celda=  ofv.GenCelda("","ut colspan=1","","1%","","","","","si")
         ' saccion = "<a class=nt href=" & chr(34) & browdir & "/" & sfolder(f)  & chr(34)
	     ' saccion = saccion & "  target=_blank ><font color=#00aa00 face=CorFlatIcon size=2 >0</font>&nbsp;&nbsp;" & REPLACE(browdir  &"/" & sfolder(f),"\","/") & "</a>"
	      saccion   = "<a class=nt href=CordscrBrow.asp?direct=" & direct & "&nomdir=" & nomdir
          saccion   = saccion & "&rutdir=" & rutdir & "&browdir=" & REPLACE(browdir  &"/" & sfolder(f),"\","/") & "   >" & "<font color=#00aa00 face=CorFlatIcon size=2 >0</font>&nbsp;&nbsp;" & REPLACE(browdir  &"/" & sfolder(f),"\","/")  & "</a>"

	      h3      = Celda
	      h1      = ofv.GenCelda("","ut","","","","","",saccion,"si")
	      h2      = ofv.GenCelda("","ut","","","","","",  "DIR","si")
	      sHTML   = sHTML & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","22","","",h3 & h1 & H2 ,"si") &formobj

	      if  xTrColor = "#FFFFFF" then
	          xTrColor = "#F9F9F9"
	      ELSE
	          xTrColor = "#FFFFFF"
     END IF

    next
   for f=0 to ubound(sfiles)-1

     xVersion = ""
     file=split(sfiles(f),"###")
       XREF="CordscrBrow.ASP?OP=DEL"
       if bSuprimir then '' and session("APLMDCLI")="BANELCO" then
      celda= BotonEliminarlnk("Confirma la eliminación de " & file(0) & "?","",xref,f,file(0))
      else

      celda=  ofv.GenCelda("","ut colspan=1","","1%","","","","","si")
      end if


     saccion = "<a class=nt href=" & chr(34) & browdir & "/" & file(0)  & chr(34)
     saccion = saccion & "  target=_blank ><font color=#00aa00 face=CorFlatIcon size=2 >&#0158;</font>&nbsp;&nbsp;" & file(0) & "</a>"
     h3      = Celda
     h1      = ofv.GenCelda("","ut","","","","","",saccion,"si")
     h2      = ofv.GenCelda("","ut","","","","","",  file(1),"si")
     sHTML   = sHTML & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","22","","",h3 & h1 & H2 ,"si") &formobj

     if  xTrColor = "#FFFFFF" then
         xTrColor = "#F9F9F9"
     ELSE
         xTrColor = "#FFFFFF"
     END IF

  Next



    sHTML = sHTML & ofv.GenFTabla()


  Agregar = ""
  Buscar  = ""

  rutdir=left(browdir,instrRev(browdir,"/",-1) - 1)

  sAccion = "Cordscr2.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir & "&OPC=NO"
  Celda   = ofv.BotonVolver(sAccion)

       sHTML = sHTML & "<script type='text/javascript'>"
        sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
        sHTML = sHTML & "  {"
        sHTML = sHTML & "  var r=confirm(disp_msg);"
        sHTML = sHTML & "  if (r==true)"
        sHTML = sHTML & "    {"
        sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
      '  sHTML = sHTML & "    alert(fr.name);"
        sHTML = sHTML & "    fr.submit();"
        sHTML = sHTML & "    }"
        sHTML = sHTML & "  else"
        sHTML = sHTML & "    {"
        sHTML = sHTML & "    alert('Se cancelo la accion!');"
        sHTML = sHTML & "    }"
        sHTML = sHTML & "  }</script>"



  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

 Function BotonEliminarlnk(titdelete,sNombre, sHref,CLAVE,ARCH)

 	             wktitdelete=titdelete '"Confirma la eliminacion de TODAS las Relaciones?"
 	             if sNombre="" then
 	              sImageELS    = "<font face='CorFlatIcon' size=3 color='#bb0000'>r</font>"
 	              xclass="ut"
 	             else
 	              sImageELS    =sNombre
 	              xclass="ttt"
 	             end if
 	             sformobj   = "FRMOBJ" & CLAVE
 	             sAccion    = sHref

 	             if  SESSION("FILETIPO") = "R" then
 	             SINPUT="<INPUT TYPE=HIDDEN NAME=ARCH VALUE=""" & replace(ARCH," ","%20") & """>"
 	             directX  ="<INPUT TYPE=HIDDEN NAME=direct VALUE=""" & replace(direct," ","%20") & """>"
 	             nomdirX  = "<INPUT TYPE=HIDDEN NAME=nomdir VALUE=""" & replace(nomdir," ","%20") & """>"
				 rutdirX  = "<INPUT TYPE=HIDDEN NAME=rutdir VALUE=""" & replace(rutdir," ","%20") & """>"
                 BrowdirX = "<INPUT TYPE=HIDDEN NAME=Browdir VALUE=""" & replace(Browdir," ","%20") & """>"
                 else
                  SINPUT="<INPUT TYPE=HIDDEN NAME=ARCH VALUE=""" & ARCH & """>"
				 directX  ="<INPUT TYPE=HIDDEN NAME=direct VALUE=""" & direct & """>"
				 nomdirX  = "<INPUT TYPE=HIDDEN NAME=nomdir VALUE=""" & nomdir & """>"
				 rutdirX  = "<INPUT TYPE=HIDDEN NAME=rutdir VALUE=""" & rutdir & """>"
                 BrowdirX = "<INPUT TYPE=HIDDEN NAME=Browdir VALUE=""" & Browdir & """>"
                 end if




 	             formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","",SINPUT & directX & nomdirX & rutdirX & BrowdirX ,"si")
 	             sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "

 	            BotonEliminarlnk=ofv.GenCelda("",xclass & " style='cursor:pointer;'  title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageELs & formobj ,"si")



End Function

%>