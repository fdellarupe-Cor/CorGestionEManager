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
  Formulario = ofv.Encabezados("VALOR","Descripcion","PROBABILIDAD","IMPACTO")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN


        strsql = "Insert into ORACTRIESGOBRUTO "
        strsql = strsql & " (valor,DESCRI,nivprob,nivimpa)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('99','99','99','99')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


  WKSETTABLA = "ORACTRIESGOBRUTO"


  Strsql = "Select * From " & WKSETTABLA & " Order By NIVPROB DESC ,VALOR DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     rs.move(CantRegMoverx)


  END IF 

  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(2),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'VALOR
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"5","dtN","")
    Celda = Celda & ofv.gencelda("","cc","","10%","","","",sInput,"si")

    'Descripcion
    Strsql = "Select * From ORRANGORIESGO ORDER BY VALOR DESC"
    sInput = ofv.GenerarCombo(sForm,RS(2).NAME,RS(2).value,Strsql,cn,0,2,"")
 ''   sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","UT","","40%","","","",sInput,"si")

    'NIVPROB
    Strsql = "Select * From ORACTPROBOCURRENCIA ORDER BY VALOR"
    sInput = ofv.GenerarCombo(sForm,RS(3).NAME,RS(3).value,Strsql,cn,0,2,"")
    Celda = Celda & ofv.gencelda("","UT BGCOLOR='#F7EEDD'","","20%","","","",sInput,"si")

    'NIVIMPA
    Strsql = "Select * From ORACTIMPACTO ORDER BY VALOR"
    sInput = ofv.GenerarCombo(sForm,RS(4).NAME,RS(4).value,Strsql,cn,0,2,"")
    Celda = Celda & ofv.gencelda("","UT BGCOLOR='#ddddff' ","","20%","","","",sInput,"si")



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
    sAccion = "RGACTRIESGOBRUTO.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " Order By VALOR DESC "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " Order By descri "    ' orden alfabetico
 ' Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "RGACTRIESGOBRUTOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>