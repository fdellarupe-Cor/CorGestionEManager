<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.menuHeader("","")

 

  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)



  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")



    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenX  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>4</FONT></B>"
    sImagenIX = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>2</FONT></B>"

    sImagenpry = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>&#0143;</FONT></B>"
    sImagenreq = "<B><FONT FACE=WEBDINGS COLOR=#006600 SIZE=3>&#0153;</FONT></B>"



  celdas = ofv.GenCelda("","TTR ","center","20%","","","","<A CLASS=NT target='Principal' HREF='PRYadmincidenciasdet.asp' >INCIDENCIAS</A>","si")
  celdas = celdas & ofv.GenCelda("","TTR ","center","20%","","","","<A CLASS=NT target='Principal' HREF='PRYasgnusr.asp' >TAREAS</A>","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","20%","","","","<A CLASS=NT target='Principal' HREF='PRYadmULTIMPUT.asp' >ULTIMAS IMPUTACIONES</A>","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","40%","","","","&nbsp;","si")

  filas = ofv.GenRow("","","","","22","","",celdas,"si")
  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")




  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Agregar)





  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML




%>