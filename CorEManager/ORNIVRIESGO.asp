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


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion","DESDE","HASTA","REQ.PLAN","REQ.INDICADOR")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN



        strsql = "Insert into ORNIVRIESGO "
        strsql = strsql & " (DESCRI,VALORDESDE,VALORHASTA,REQPLAN,REQINDICADOR)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('0','0','0','NO','NO')"
        xxx=SqlExLog("INSERT",strsql,CN,"ORNIVRIESGO","","","PAR")
  '''response.write strsql & "<br>"
       ' call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF

  WKSETTABLA = "ORNIVRIESGO"

  Strsql = "Select * From " & WKSETTABLA & " Order By VALORDESDE, VALORHASTA"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Descripcion
    Strsql = "Select * From ORRANGORIESGO ORDER BY VALOR DESC"
    sInput = ofv.GenerarCombo(sForm,RS(1).NAME,RS(1).value,Strsql,cn,0,2,"")
  '''  sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","UT","","30%","","","",sInput,"si")

    'VALORDESDE
    sInput = ofv.GenerarInput(rs(2).name,REPLACE(CSTR(rs(2) & ""),",","."),"text",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","ut","","15%","","","",sInput,"si")

    'VALORHASTA
    sInput = ofv.GenerarInput(rs(3).name,REPLACE(CSTR(rs(3) & ""),",","."),"text",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","UT","","15%","","","",sInput,"si")

    'REQUIEREPLAN
    scomK = "SI" & chr(9) & "NO"
    scomD = "SI" & chr(9) & "NO"
    sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")
   ''' sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"2","dt","")
    Celda = Celda & ofv.gencelda("","UT","","10%","","","",sInput,"si")

     'REQUIER INDICADOR
	 scomK = "SI" & chr(9) & "NO"
	 scomD = "SI" & chr(9) & "NO"
	 sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),scomK,scomD," ")

     Celda = Celda & ofv.gencelda("","UT","","10%","","","",sInput,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORNIVRIESGO.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " Order By VALORDESDE "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " Order By descri "    ' orden alfabetico
 ' Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORNIVRIESGOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>