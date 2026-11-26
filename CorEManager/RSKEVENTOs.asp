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
  WKRSKSIT = request.FORM("RSKSIT") & ""

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))


  IF WKRSKSIT = "" THEN 

     Strsql = "Select * From RSKSITUACION ORDER BY DESCRI "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then 

        WKRSKSIT = RS(0)

     ELSE 

        WKRSKSIT = "0"

     END IF


  END IF

    sForm = "rskeven"

    Sql1 = "Select * From RSKSITUACION  Order By descri"
    sInput = ofv.GenerarCombo(sForm,"RSKSIT",WKRSKSIT,Sql1,cn,0,1,"")

    Celda = ofv.GenCelda("","TTt","","","","","","SITUACION DE RIESGO:","si")
    Celda = Celda & ofv.gencelda("","UT","","","95%","","",sInput,"si")


    sAccion = "RSKEVENTOs.ASP"

    Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")

    Fila = ofv.genrow("","","","","","","",Formulario,"si")

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si") & "<BR>"




  'Encabezados
  Fila = ofv.Encabezados("CATEGORIA","EVENTOS")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar




  Strsql = "Select * From RSKCATEGO WHERE SITUACION = '" & WKRSKSIT & "' Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi< CantRegMostrarx)
    Formulario = ""
    sForm = "form" & posi

    'Eliminar

      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"50","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")



    sImagenP = "&#0232;"
    sFuncAnc = " onmouseover=""this.className='ft2o';"" onmouseout=""this.className='ft2n';"" "
    sinput = "<a class='ft2n' " & sFuncAnc & " href='RSKEVENTODET.asp?CATEGO=" & RS(0)
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
  Buscar = ""

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>