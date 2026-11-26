<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

    Fila = ""

    sInput = ofv.GenerarInput("","CONSULTAS","submit","","10","bt","")
    Celda = ofv.GenCelda("","","","10%","","","",sInput,"si")

    sAccion = "DepNoveCons.asp"
    Formulario = ofv.GenForm("Boton1","",sAccion,"","post","",Celda,"si")



    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

    sInput = ofv.GenerarInput("","DEPURACION","submit","","10","bt","")
    Celda = ofv.GenCelda("","","","10%","","","",sInput,"si")

    sAccion = "DepNoveDepu.asp"
    Formulario = ofv.GenForm("Boton2","",sAccion,"","post","",Celda,"si")



    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

    sInput = ofv.GenerarInput("","LOG DEPURACION","submit","","10","bt","")
    Celda = ofv.GenCelda("","","","10%","","","",sInput,"si")

    sAccion = "DepNoveLog.asp"
    Formulario = ofv.GenForm("Boton3","",sAccion,"","post","",Celda,"si")



    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")



  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si")



  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>