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

  sHTML = ofv.FormHeader("AGREGAR PERFILES DE COMPONENTES")


  sForm = "Agregarb"
  sFormX = "AgregarX"

    sInput = "AGREGAR " & session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","Perfil","si")
     sInput = ofv.GenerarInput(rs(0).name,"","text",sFormX,"40","dt","")
     sInput = REPLACE(sInput,"{ document." & sFormX  & ".submit(); return true;}","{ if (this.value.length > 10) {alert('El texto es demasiado extenso'); return false;} else { document." & sFormX & ".submit(); return true;}}")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
     sInput = ofv.GenerarInput(rs(1).name,"","agregar","","90","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sInput   = ofv.GenerarInput(rs(2).name,"1","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput(rs(3).name,"1","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput(rs(4).name,"1","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput(rs(5).name,"1","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput(rs(6).name,"0","HIDDEN","","1","dt","")

  sInput  = sInput &  ofv.GenerarInput("CMP06","1","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP07","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP08","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP09","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP10","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP11","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP12","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP13","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP14","0","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMP15","0","HIDDEN","","1","dt","")

  sInput  = sInput &  ofv.GenerarInput("CMPD01","PROPIEDADES","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD02","ESTRUCTURA","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD03","DETALLE","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD04","IMAGENES","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD05","CARPETA","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD06","ARCHIVOS","HIDDEN","","1","dt","")

  sInput  = sInput &  ofv.GenerarInput("CMPD07","SD","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD08","SD","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD09","SD","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD10","SD","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD11","SD","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD12","SD","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD13","SD","HIDDEN","","1","dt","")
 sInput  = sInput &  ofv.GenerarInput("CMPD14","SD","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPD15","SD","HIDDEN","","1","dt","")
'
  sInput  = sInput &  ofv.GenerarInput("CMPO01","3","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO02","4","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO03","5","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO04","1","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO05","6","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO06","2","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO07","99","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO08","99","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO09","99","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO10","99","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO11","99","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO12","99","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO13","99","HIDDEN","","1","dt","")

 sInput  = sInput &  ofv.GenerarInput("CMPO14","99","HIDDEN","","1","dt","")
  sInput  = sInput &  ofv.GenerarInput("CMPO15","99","HIDDEN","","1","dt","")



  Celda   = ofv.gencelda("","vt colspan=2","","","","","",sInput,"si")
  Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")



  'sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
  'sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0 "
  'sFuncion = sFuncion & " || this." & rs(0).name & ".value==' ') "
  'sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;}" & chr(34)
  sFuncion =  chr(34) & " var fr=document.Agregarb; if (fr." & rs(0).name
  sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.length==0 "
  sFuncion = sFuncion & "|| fr." & rs(0).name & ".value==' ') "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=../PerfilesdeComponentes.asp?CantRegMover=" & Regisvuelta
  Formulario = Formulario & ofv.genform("Agregarb","",sAccion,"","post","",Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")

  'Aceptar y Cancelar
  sAccion = "PerfilesdeComponentes.asp?CantRegMover=" & Regisvuelta
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

response.write sHTML

%>