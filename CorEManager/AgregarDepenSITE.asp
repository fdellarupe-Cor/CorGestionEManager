<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  NOM = request.querystring("NOM")
  NOMU = request.querystring("NOM")
  NOM = ofv.convertircar(NOM,"*","&nbsp;")

  ID = request.querystring("ID")
  Tabla = request.querystring("Tabla")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sHTML = ofv.FormHeader("AGREGAR DEPENDENCIA SITE/SUCURSAL")

  f1f = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  f1f = f1f & ".value.length==0) {alert('Los campos no deben ser nulos');return false;}" & chr(34)
  f1a = "asp/Agregar.asp?Tabla=" & Tabla
  f1a = f1a & "&volver=../SectorDepenSITES.asp"
  sHTML = sHTML & ofv.GenForm("Agregarbe","",f1a,"","post",f1f,"","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados
  h1 = ofv.GenCelda("","tt","","2% colspan=2","","","","Sector Dependiente: " & NOM,"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  Sql1 = "Select * From SITES where SITEun = " & session("sun") & " order by descri"
  t1 = ofv.GenCelda("","","","2% colspan=2","","","",ofv.GenerarInput(rs(1).name,ID,"hidden","","","",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t1,"si")
  'Dependencias
  t2 = ofv.GenCelda("","tu","","2%","","","","SITE/SUCURSAL:","si")
  t2 = t2 & ofv.GenCelda("","","","","","","",ofv.GenerarCombox3(sql1,cn,0,1,rs(2).name),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t2,"si")
  sHTML = sHTML & ofv.GenFTabla()

  'Botones de aceptar/cancelar
  sAccion = "SectorDepenSITES.asp"
  sHTML = sHTML & ofv.AgregarBoton("Agregarbe",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>