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

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  if Session("CorApli") = "COREM20" then
     wkaplic = "EM"
  elseif Session("CorApli") = "COREP20" then
     wkaplic = "PL"
  else
     wkaplic = "OR"
  end if


  sHTML = ofv.FormHeader("AGREGAR TIPOS DE SUBTEMA")

  sForm = "Agregarb"
  sFormX = "AgregarX"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","Tipo","si")
     sInput = ofv.GenerarInput(rs(0).name,"","text",sFormX,"40","dt","")
     sInput = REPLACE(sInput,"{ document." & sFormX  & ".submit(); return true;}","{ if (this.value.length > 10) {alert('El texto es demasiado extenso'); return false;} else { document." & sFormX & ".submit(); return true;}}")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","90","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


     'Celda  = ofv.gencelda("","tt","","","","","","Icono","si")
     'sImagenV = ucase("ICorAbierto.jpg")
     'sImagenC = rs(15).name
     'sInput = ofv.GenerarCombo333(ofv.Getfiles("Iconos"),sImagenC,sImagenV,"Iconos\","imagen")
     'Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    'Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Nivel","si")
     sInput = "B"
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
 FILT="C"
  IF SESSION ("BASE") ="K" THEN
  FILT="K"
  else
  FILT="C"
  END IF
  sInputICONO   = ofv.GenerarInput("ICONO",FILT,"HIDDEN","","1","dt","")

  sInput   = ofv.GenerarInput(rs(4).name,"B","HIDDEN","","1","dt","")
  For i=0 to 9
      sInput  = sInput &  ofv.GenerarInput(rs(5 + i).name,"1","HIDDEN","","1","dt","")
  next
  sInput  = sinput & ofv.GenerarInput(rs(16).name,wkaplic,"hidden","","1","dt","")
  Celda   = ofv.gencelda("","vt colspan=2","","","","","",sInput  & sInputICONO,"si")
  Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  'sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
  'sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0 "
  'sFuncion = sFuncion & "|| this." & rs(0).name & ".value==' ') "
  'sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;}"
  'sFuncion = sFuncion & "if (this." & rs(0).name
  'sFuncion = sFuncion & ".value.length>3 "
  'sFuncion = sFuncion & ") "
  'sFuncion = sFuncion & "{alert('Tipo no debe superar los tres caracteres');return false;}" & chr(34)

   sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(0).name
  sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.length==0 "
  sFuncion = sFuncion & "|| fr." & rs(0).name & ".value==' ') "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../TiposdeSubtema.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.genform("Agregarb","",sAccion,"","post","",Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")

  'Aceptar y Cancelar
  sAccion = "TiposdeSubtema.asp?CantRegMover=" & Regisvuelta
  'sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self") & chr(13)
   'sAccion = "PerfilesdeEsquemaNormativo.asp?CantRegMover=" & Regisvuelta
  zboton=ofv.AgregarBoton("Agregarb",sAccion,"") & chr(13)
  zboton=replace(zboton,"'document.Agregarb.submit();'",sFuncion)
  sHTML = sHTML & zboton

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

set ofv = nothing

response.write sHTML

%>