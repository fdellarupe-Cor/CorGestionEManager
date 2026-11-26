<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  nID = request.querystring("ID")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")

  set cn = ofv.conectar(ofv.strconn2)
  'response.write Tabla
  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("Agregar Descripcion de Control")
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

    f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo & "&Control=" & nID
  sHTML = sHTML & ofv.GenForm("AgregarControlb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezado Descripcion
  h1 = ofv.GenCelda("","tt","","2%","","","","Sec","si")
  h1 = h1 & ofv.GenCelda("","tt","","98%","","","","Descripcion","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

      sql = "SELECT max(" & tabla & ".secuencia) FROM " & tabla & " where "
      sql = sql & "control = " & nID
      'response.write sql
      set rsSecuencia = ofv.crearconsultaEx(Sql,cn,1,volver)
      nSecuencia = 10
      if not rsSecuencia.eof then
         if not isnull(rsSecuencia(0)) then nSecuencia = cint(rsSecuencia(0)) + 10
      end if
      call ofv.cerrarconsulta(rsSecuencia)

  sInput = ofv.GenerarInput("secuencia",nSecuencia,"agregar","","10","dtn","SI")
  sInput1 = ofv.GenerarInput("Modi","1","hidden","","10","","")
  sInput2 = ofv.GenerarInput("TipModi","A","hidden","","10","","")
  t1 = ofv.GenCelda("","","","","","","",sInput & sinput1 & sinput2,"si")

  esttxt = " STYLE=font-family:tahoma;color:#225588;font-size:7pt;background-color:#f7eedd; "

  'Texto Control
  t1 = t1 & ofv.GenCelda("","","","","","","","<font class=at><textarea name=descri rows=5 " & esttxt & " cols=120></textarea></font>","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t1,"si")
  sHTML = sHTML & ofv.GenFTabla()
 ' response.write right(volver,len(volver)-3)
  'Botones Aceptar y Cancelar
  sAccion = ofv.convertircar(right(volver,len(volver)-3),"_","&") & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgregarControlb",sAccion,"_self")


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
   sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>