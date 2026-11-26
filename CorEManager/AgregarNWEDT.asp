<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo & " Order By consecuen desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    nSecuencia = clng(rs("consecuen")) + 10
  else
    nSecuencia = 10
  end if

  sHTML = ofv.FormHeader("Agregar Indice")

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

  sForm = "Agregarb"

    sInput = "AGREGAR DATOS"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC


     Celda  = ofv.gencelda("","tt","","20%","","","","SECUENCIA","si")
     sInput = ofv.GenerarInput(rs(4).name,nSecuencia,"Agregar","","5","dtn","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","CAPITULO","si")
     sInput = ofv.GenerarInput(rs(3).name,"","Agregar","","90","dt","") & ofv.GenerarInput(rs(5).name,XNNWCOD,"hidden","","1","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sFuncion = ""
  sAccion  = "asp/AgregarEN.asp?Tabla=" & Tabla
  sAccion  = sAccion & "&volver=" & volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


  sHTML = sHTML & "<center>"
  Fila    = ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Fila,"si")
  sHTML   = sHTML & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")
sHTML = sHTML & "</center>"


  'Botones Aceptar y Cancelar
  sAccion = "edNWEDT.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML   = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")


  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
  Filann = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%   ","","0","","1","0","","",Filann,"si")



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"
  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>