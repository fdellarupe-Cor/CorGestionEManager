<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Fila = ofv.Encabezados("SITUACION DE RIESGO","CATEGORIAS")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar




  Strsql = "Select * From RSKSITUACION  Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi< CantRegMostrarx)
    Formulario = ""
    sForm = "form" & posi

    'Eliminar

      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"50","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")



    sImagenP = "&#0232;"
    sFuncAnc = " onmouseover=""this.className='ft2o';"" onmouseout=""this.className='ft2n';"" "
    sinput = "<a class='ft2n' " & sFuncAnc & " href='RSKCATEGODET.asp?SITUACION=" & RS(0)
    sinput = sinput  & "' title='CATEGORIAS de " & rs(1) & "' >" & sImagenP & "</a>" 
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    sAccion = ""

    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

    posi=posi+1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si")

  Agregar = ""


  str1 = "Select descri From RSKSITUACION Order By descri "    ' orden pagina
  str2 = "Select descri From RSKSITUACION Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>