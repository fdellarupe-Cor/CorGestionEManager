<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.FormHeader(session("FormN"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sPath = Request.ServerVariables("PATH_INFO")
  Xruta = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT") & left(sPath,instrRev(sPath,"/",-1))


  Xfer =  Xruta
  XferTarget1 = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "componentes"
  nomdir1 = "Imagenes&nbsp;de&nbsp;Componentes"
  XferTarget2 = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "estructura"
  nomdir2 = "Imagenes&nbsp;de&nbsp;Contenidos"
  XferTarget3 = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "archivos"
  nomdir3 = "Documentos&nbsp;Asociados"
  XferTarget4 = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "arccomponentes"
  nomdir4 = "Archivos&nbsp;de&nbsp;Componentes"

   sicK = "<font face='corflaticon' size=4 color='" & session("cmfmenu") & "' >&#0158;</font>"


   'Encabezados
    Celda = ofv.GenCelda("","ttt colspan=2","","","","","","Directorios de imagenes y archivos","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccioni = "<a href=Cordscr2.asp?direct=" & "componentes" & "&nomdir=" & nomdir1 & "&rutdir=" & XferTarget1 & "  >" & sicK & "</a>"
      sAccion  = "<a class=nt href=Cordscr2.asp?direct=" & "componentes" & "&nomdir=" & nomdir1 & "&rutdir=" & XferTarget1 & "  ><big>Imagenes de Componentes</big></a>"
    end if   

    Celda = ofv.GenCelda("","uts","center","5%","","","",saccioni,"si")
    Celda = Celda & ofv.GenCelda("","uti","","95%","","","",saccion,"si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccioni = "<a href=Cordscr2.asp?direct=" & "estructura" & "&nomdir=" & nomdir2 & "&rutdir=" & XferTarget2 & " >" & sicK & "</a>"
      sAccion  = "<a class=nt href=Cordscr2.asp?direct=" & "estructura" & "&nomdir=" & nomdir2 & "&rutdir=" & XferTarget2 & " ><big>Imagenes de Contenidos</big></a>"
    end if  

    Celda = ofv.GenCelda("","uts","center","5%","","","",saccioni,"si")
    Celda = Celda & ofv.GenCelda("","uti","","95%","","","",saccion,"si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccioni = "<a href=Cordscr2.asp?direct=" & "archivos" & "&nomdir=" & nomdir3 & "&rutdir=" & XferTarget3 & " >" & sicK & "</a>"
      sAccion  = "<a class=nt href=Cordscr2.asp?direct=" & "archivos" & "&nomdir=" & nomdir3 & "&rutdir=" & XferTarget3 & " ><big>Documentos Asociados</big></a>"
    end if  

    Celda = ofv.GenCelda("","uts","center","5%","","","",saccioni,"si")
    Celda = Celda & ofv.GenCelda("","uti","","95%","","","",saccion,"si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccioni = "<a href=Cordscr2.asp?direct=" & "arccomponentes" & "&nomdir=" & nomdir4 & "&rutdir=" & XferTarget4 & " >" & sicK & "</a>"
      sAccion  = "<a class=nt href=Cordscr2.asp?direct=" & "arccomponentes" & "&nomdir=" & nomdir4 & "&rutdir=" & XferTarget4 & " ><big>Archivos de Componentes</big></a>"
    end if  

    Celda = ofv.GenCelda("","uts","center","5%","","","",saccioni,"si")
    Celda = Celda & ofv.GenCelda("","uti","","95%","","","",saccion,"si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




  sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","","","50%","","0","","2","5","","",Fila,"si")



  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","",Buscar,celda)

 

  sHTML = sHTML & "</center></body></html>"
end if

Response.Write sHTML

%>