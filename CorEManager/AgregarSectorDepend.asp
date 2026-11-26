<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  Regisvuelta = request.querystring("Regisvuelta")
  NOM = request.querystring("NOM")
  NOM = ofv.convertircar(NOM,"*","&nbsp;")
  ID = request.querystring("ID")
  Tabla = request.querystring("Tabla")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)  

  sHTML = ofv.FormHeader("AGREGAR DEPENDENCIA")

  f1f = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  f1f = f1f & ".value.length==0) {alert('Los campos no deben ser nulos');return false;}" & chr(34)
  f1a = "asp/Agregar.asp?Tabla=" & Tabla
  f1a = f1a & "&volver=../SectorDependencias.asp?CantRegMover=" & Regisvuelta & "&ID=" & ID 
  sHTML = sHTML & ofv.GenForm("Agregarb","",f1a,"","post",f1f,"","no")

  sHTML = sHTML & ofv.GenTabla("panel","aaa","","60%","",0,0,0,0,"","","","no")

  'Encabezados
  h1 = ofv.GenCelda("","ttT","","2% colspan=2","","","","Sector Dependiente: " & NOM,"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  Sql1 = "Select * From Eco where ecoun = " & session("sun") & " order by descri"
  t1 = ofv.GenCelda("","","","2% colspan=2","","","",ofv.GenerarInput(rs(2).name,ID,"hidden","","","",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t1,"si")
  'Dependencias
  t2 = ofv.GenCelda("","TT","","20%","","","","Sector&nbsp;Dependencia:","si")
  t2 = t2 & ofv.GenCelda("","vti","","80%","","","",ofv.GenerarCombox3(sql1,cn,0,1,rs(1).name),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t2,"si")
  sHTML = sHTML & ofv.GenFTabla()
  
  'Botones de aceptar/cancelar
  sAccion = "SectorDependencias.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
 
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>