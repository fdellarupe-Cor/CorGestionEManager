<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados

  Formulario = ofv.Encabezados("NIVEL","DESCRI","FUNCIONES&nbsp;GEN.","COLOR","DET&nbsp;COLOR")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
  session("MODORIGEN")="LISTAPUESTOS"
  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From NivelPUESTO Order By NIVELP"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    sInput = rs(0)

    'Eliminar
    if bSuprimir and sinput > "0" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=NivelPUESTO&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'ID
  '  sInput = rs(0)
  '  Celda = Celda & ofv.GenCelda("","cc","right","50","","","",sInput,"si")

    'NIVELP
    IF BMODIFICAR THEN
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"10","dtn","")
    ELSE
    sInput=rs(1)
    END IF

    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

     'DESCRI
      IF BMODIFICAR THEN
     sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"100","dt","")
     ELSE
     sInput=RS(2)
     END IF
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


    ssino = rs(3) & ""

 '  if ssino = "" then ssino = "N"
'
 '  scomK = "S" & chr(9) & "N"
 '  scomD = "SI" & chr(9) & "NO"
'
'
  '  if bModificar then
  ''      sInput = ofv.gencomboFX(sForm,rs(3).name,ssino,scomK,scomD," ")
 '   else
 ''       sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
 '   end if
'
   ''''' sInput = ofv.combosino(sForm,ssino,rs(2).name)
  ' Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Celda = Celda & ofv.gencelda("","","","","","","","","si")

   'Propiedades
       'sImagenP = "&#0232;"
       'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
       sinput = "<a href='sectorperFilESfun2.asp?solapa=FUN&PERFIL=" & rs(1)
       'sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx

       sinput = sinput  & "' title='Propiedades de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"

       if bModificar then
       Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")
       ELSE
       Celda = Celda & ofv.gencelda("","ft","center","","15%","","","&nbsp;","si")
       END IF

     sql3   = "Select * From SYSCOLOR Order By DESCRI"
     IF BMODIFICAR THEN
     sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,sql3,cn,2,1,"")
     ELSE
     sInput="&nbsp;"
     END IF

     celdaz=ofv.GenCelda("","aa style='border-width: 1px;border-style:solid;border-color:#ffffff;' bgcolor=" & rs(4),"","","75%","","","&nbsp;","si")
     celda  = celda & ofv.GenCelda("","","","","","","",sInput,"si") & celdaz

    Fila = ofv.GenRow("","PDC style='cursor:pointer; '","","","","","",Celda,"si")
                          'style='border-style:solid; border-color: #000000 border-width: 1px'
    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=NIVELPUESTO&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarNivelPUESTO.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=NivelPUESTO"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From NivelPUESTO Order By NIVELP "    ' orden pagina
  str2 = "Select descri From NivelPUESTO Order By NIVELP "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sinput = "<a  href=NivelPUESTOCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  sAccion = "NivelePUESTOCons.asp?" & request.querystring
''Celda = ofv.BotonConsultar(sAccion)
  Reportes=BtnReportes(16,"","xform",sAccion)

''sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,"no",Reportes)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>