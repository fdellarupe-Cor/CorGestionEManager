<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  volver = request.querystring("volver")
  sUser = request.querystring("Usuario")

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader("Asignar Grupo")

    f1a = "asp/AsignarGrupo.asp?Usuario=" & sUser
  sHTML = sHTML & ofv.GenForm("Agregarb","",f1a,"info","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezado Grupo
  h1 = ofv.GenCelda("","tt","","2%","","","","Grupo","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Combo Grupo
  strsql = "Select * From Grupos Order By Descripcion"
  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarCombox2(strsql,cn,0,1,"Grupo"),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t1,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = "Usuarios.asp"
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>