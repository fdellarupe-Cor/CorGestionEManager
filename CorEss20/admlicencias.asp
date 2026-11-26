<%@LCID = 11274%>
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else


  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sopcion = request.querystring("opcion")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



    sImagen4 = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"



  sHTML = ofv.FormHeader(session("FormN"))

''  sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (document.getElementById('odiv')) { document.getElementById('odiv').style.top=document.body.clientHeight - 30; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)

  'Encabezados
  Formulario = ofv.Encabezados("licencia","empresa")


  Strsql = "Select OBJID,CODLICENCIA,CODEMPRESA From corconsulta order by codempresa"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi


      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")

    'Descripcion
    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","vt","left","20%","","","",sInput ,"si")


     strsql = "Select * From clienteplan Order By cliente"
     sInput = ofv.GenerarCombo(sForm,rs(2).name,rs(2),strsql,cn,0,1,"")
    'Descripcion
    'sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"80","dt","")
    Celda = Celda & ofv.gencelda("","","","50%","","","",sInput,"si")




    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
     SACCION = "asp/actualizar.asp?id="
     SACCION = SACCION & rs(0) & "&tabla=corconsulta&volver=../admlicencias.asp"
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  ftab = ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","0","0","","",filas,"si")



 '' sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")

  'Botonera
   Agregar = ""
   Buscar = ""






  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

 end if




Response.Write sHTML

%>