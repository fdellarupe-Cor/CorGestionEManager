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
  nEsquema = request.querystring("Esquema")
  sTipo = request.querystring("Tipobjeto")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("Agregar Cuenta Contables")
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------
    f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo & "&Esquema=" & nEsquema
  sHTML = sHTML & ofv.GenForm("AgregarCuentab","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Cuenta Contable y Sentido
  h1 = ofv.GenCelda("","ttR","","80%","","","","Cuenta Contable","si")
  h1 = h1 & ofv.GenCelda("","ttR","","20%","","","","Sentido","si")
  sHTML = sHTML & ofv.GenRow("","","","","25","","",h1,"si")

  'Combos Cuentas Contables
  Strsql="Select * From Objcuenta where estado=3 Order By descri"
  Array1 = ofv.GenerarArray(strsql,cn,0)
  Array2 = ofv.GenerarArray(strsql,cn,1)
  Array3 = ofv.GenerarArray(strsql,cn,2)

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""
     scomS   = "0"

     For n= 1 To Ubound(Array1)
         if   Array1(n) <> "" then
              if  scomS   = "0" then      scomS   = ofv.convertircar(Array1(n)," ","_")
              scomK = scomK & SCOMSEP & ofv.convertircar(Array1(n)," ","_")
              scomD = scomD & SCOMSEP & Array3(n) & " - " & Array2(n)
              SCOMSEP = chr(9)
         end if
     next



    t1 = ofv.gencomboFX(sForm,"codcuenta",scomS,scomK,scomD,"no")


  sInput1 = ofv.GenerarInput("Modi","1","hidden","","10","","")
  sInput2 = ofv.GenerarInput("TipModi","A","hidden","","10","","")
  t1 = ofv.GenCelda("","","","","","","",t1 & sinput1 & sinput2,"si")

' t2 = "<font class=at><select name=sentido>"
'  t2 = t2 & "<option value=Debe>Debe<option value=Haber>Haber"
'  t2 = t2 & "</select></font>"

      'Sentido
    scomK = "Debe" & chr(9) & "Haber"
    scomD = "DEBE" & chr(9) & "HABER"

    t2 = ofv.gencomboFX(sForm,"sentido","Debe",scomK,scomD,"no")




  t2 = ofv.GenCelda("","","","","","","",t2,"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2,"si")
  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
 '' sAccion = "edEsquemas.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Esquema=" & nEsquema
  sAccion = replace(replace(Volver,"../",""),"_","&") & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo  & "&Esquema=" & nEsquema
  sHTML = sHTML & ofv.AgregarBoton("AgregarCuentab",sAccion,"_self")

  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
  call ofv.cerrarconsulta(rs)
  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>