<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Tabla = request.querystring("Tabla")
  vuelta = request.querystring("vuelta")
  Regisvuelta = request.querystring("Regisvuelta")
  INDICE=request.querystring("Indice")
  if request.querystring("Opcion") <> "" then
     Opcion = request.querystring("Opcion")
     volver = request.querystring("volver")
  else
     Opcion = "no"
     volver = "no"
  end if

  set cn = ofv.conectar(ofv.strconn2)



  strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR ACCIONES")


  IF Indice <>"NO" THEN
  '----------------------------------------------

     shtml = shtml & CorPnlFnc("N","")

     sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

END IF

  sForm = "Agregarb"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(1).name,"ACCION","agregar","","100","dt","")
     sInput2 = ofv.GenerarInput("ESTADO","3","HIDDEN","","10","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput &sInput2 ,"si")
     Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","tipo de accion","si")
     Strsql = "Select * From Modo Order By descri"
     sInput = ofv.GenerarCombo(sForm,rs(2).name,"OP",Strsql,cn,0,1,"no")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
     Fila   = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0 || this." & rs(2).name & ".value.length==0) "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla
  if Opcion = "no" then
    sAccion = sAccion & "&volver=../Acciones.asp?CantRegMover=" & Regisvuelta
  else
    sAccion = sAccion & "&volver=" & volver
  end if
  sAccion = sAccion & "&Campo=1"
  
  tablaN =  ofv.GenTabla("","aaa","","","","0","","0","0","","",fila,"si")  
  Formulario = ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,tablaN,"no")
  'Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

  'sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
  sHTML = sHTML & Formulario

  'Aceptar y Cancelar
  if Opcion = "no" then
     sAccion = "Acciones.asp?CantRegMover=" & Regisvuelta
     sTargetx = ""
  else
     sAccion = ofv.convertircar(volver,"_","&")
     sAccion = mid(sAccion,4,len(sAccion) - 3)
     sTargetx = "_self"
  end if
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,sTargetx)


  Agregar = ""
  Buscar  = ""
  Celda   = ""

 xxxx=ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  IF Indice <>"NO" THEN
  '--------------------------------------------------------------------------------
     sHTML = sHTML & "</span></div>"

     Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

    'Estructura del Indice

     sHTML = sHTML & CorPnlIndx("N","") & xxxx
  ELSE
    sHTML = sHTML & xxxx

  '--------------------------------------------------------------------------------
  END IF

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>