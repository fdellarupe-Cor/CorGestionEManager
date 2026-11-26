<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

  nCodigo          = request.querystring("Codcomp")
  sTipo            = request.querystring("Tipocomp")
  sKstr            = request.querystring("Kstr")
  CantRegAMover    = cdbl(request.querystring("CantRegAMover"))
  CantRegAMostrar  = 10

  if  sKstr <> "" then
      session("Kstr") = sKstr
  else
      sKstr = session("Kstr")
  end if 

  vuelta           = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  'Encabezado Estructura
  Celda = ofv.GenCelda("","ttt","","2%","","","","Estructura:","si")

  Strsql = "Select * From Compestr Where codcomp = " & nCodigo & " and objid = " & sKstr
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      nID    = rs(0)
      sInput = rs(3)
      Celda  = Celda & ofv.GenCelda("","ttt","","","","","",sInput,"si")

  else
      Celda  = Celda & ofv.GenCelda("","ttt","","90%","","","","&nbsp;sin definir","si")
  end if

  Fila  = ofv.GenRow("","","","","25","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Fila,"si")

  call ofv.cerrarconsulta(rs)

 '' sHTML = sHTML & "<hr color=#dddddd size=2>"

  'Encabezados
  fila = ofv.Encabezados("Sec","Descripcion","Dato/Columna")
  posi       = 0

  if  nID <> "" then
      Strsql  = "Select * From Compdet Where codestr=" & nID & " order by datoformul, objid"
      set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if  not rs2.eof then rs2.move(CantRegAMover)
      do  until ((rs2.eof) or (posi = CantRegAMostrar))
          sForm = "edEstr" & rs2(0)
          posi  = posi + 1
      
           sAccion = ""
      if  bModificar then
          sAccion = "asp/ActualizarC.asp?ID=" & rs2(0) & "&Tabla=Compdet&volver=" & vuelta
          sAccion = sAccion & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
      end if

       'Eliminar
      if  bModificar then
          sHref = "asp/EliminarC.asp?ID=" & rs2(0) & "&Tabla=Compdet&volver=" & vuelta
          sHref = sHref & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo & "&Eliminar=NSNC"
          Celda = ofv.BotonEliminar(rs2(4),sHref)
      else
          Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

      'Secuencia
      if  bModificar then
          sInput = ofv.GenerarInput(rs2(6).name,rs2(6),"text",sForm & "X1","2","dtn","SI")
          sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")
      else
          sInput = ofv.GenerarInput("",rs2(6),"readonly","","2","dtn","")
      end if
      Celda = Celda & ofv.GenCelda("","ut","","1%","","","",sInput,"si")

      'Descripcion
      if  bModificar then
          sInput = ofv.GenerarInput(rs2(4).name,rs2(4),"text",sForm & "X2","40","dt","")
          sInput = ofv.GenForm(sForm & "X2","",sAccion,"","post","",sInput,"si")
      else
          sInput = ofv.GenerarInput("",rs2(4),"readonly","","40","dt","")
      end if
      Celda = Celda & ofv.GenCelda("","ut","","30%","","","",sInput,"si")

      'Tipo de Dato/Columna
      if  bModificar then
          sInput = ofv.GenerarInput(rs2(5).name,rs2(5),"text",sForm & "X3","70","dt","")
          sInput = ofv.GenForm(sForm & "X3","",sAccion,"","post","",sInput,"si")
      else
          sInput = ofv.GenerarInput("",rs2(5),"readonly","","70","dt","")
      end if
      Celda = Celda & ofv.GenCelda("","ut","","70%","","","",sInput,"si")

      Celda = Celda & ofv.GenCelda("","","","","","","","","si")
      Fila  = fila & ofv.GenRow("","","","","","","",Celda,"si")

      'Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

      rs2.movenext
    loop
    sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",fila,"si")

    'Agregar
    Agregar = ""
 '   if bModificar then
 '     sInput  = ofv.GenerarInput("","Agregar","submit","","10","bt","")
 '     Celda   = ofv.GenCelda("","","","","","","",sInput,"si")
      sAccion = "AgregarEstrKompDet.asp?Tabla=Compdet&volver=" & vuelta & "&Codcomp="
      sAccion = sAccion & nCodigo & "&Tipocomp=" & sTipo & "&ID=" & nID
 '     Agregar = ofv.GenForm("Agregarb","",sAccion,"","post","",Celda,"si")
 '   end if

  if bModificar then
      sAccion = "AgregarEstrKompDet.asp?Tabla=Compdet&volver=" & vuelta & "&Codcomp="
      sAccion = sAccion & nCodigo & "&Tipocomp=" & sTipo & "&ID=" & nID
      Agregar = ofv.BotonAgregar(sAccion)
  end if


    sAccion = "edEstrKomp.asp?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo 
    Celda   = ofv.BotonVolver(sAccion)

    wwwbarra = ofv.botoneraESSX("8",rs2,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no",Agregar,"no",Celda)
    call ofv.cerrarconsulta(rs2)
  end if

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

    sHTML = sHTML & wwwbarra

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>