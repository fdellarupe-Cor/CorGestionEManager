<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sTipo = request.querystring("Tipo")
  sNombre = request.querystring("Nombre")
  sNombre2 = ofv.convertircar(sNombre,"_"," ")
  sElemento = request.querystring("Elemento")
  n = instr(sElemento,"-")
  nID = mid(sElemento,n+1)

  sTitulo = "Agregar " & sNombre2
  sTitulo = UCase(sTitulo)
  sHTML = ofv.FormHeader(sTitulo)
  volver= ofv.vueltaasp2("../")



    f1a = "asp/AgregarEsquema.asp?Tipo=" & sTipo
    f1a = f1a & "&Nombre=" & sNombre & "&Elemento=" & sElemento
  sHTML = sHTML & ofv.GenForm("Agregarb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","aaa","","80%","",0,0,0,0,"","","","no")

  'Encabezado Descripcion
   t1 = ofv.GenCelda("","tt","","2%","","","","Descripcion","si")
  'Texto Descripcion
   sinput =  ofv.GenerarInput("Descri","","agregar","","60","dt","")
   t1 = t1 & ofv.GenCelda("","","","","","","",sinput,"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T1,"si")

   sinput =  ofv.GenerarInput("OBJUN",SESSION("UN"),"HIDDEN","","1","dt","")
   sinput = sinput & ofv.GenerarInput("aplic",wkaplic,"HIDDEN","","1","dt","")
   sinput = sinput & ofv.GenerarInput("prioridad","1","HIDDEN","","1","dt","")

   t2 = ofv.GenCelda("","","","","","","","&nbsp;" &  sinput ,"si")
   t2 = t2 & ofv.GenCelda("","","","","","","","&nbsp;","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T2,"si")

  'Encabezado Identificacion
   t3 = ofv.GenCelda("","tt","","2%","","","","Identificacion","si")
  'Texto Identificacion
   sinput =  ofv.GenerarInput("Codiden","","agregar","","40","dt","")
   t3 = t3 & ofv.GenCelda("","","","","","","",sinput,"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T3,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = "Objtree.asp?Nombre=" & sNombre & "&Tipo=Subtemas"
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"no")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>