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
  Formulario = ofv.Encabezados("Id","Descripcion","ESTADO")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Frecuencia order by DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar

    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")

    'id
    sInput = posi
    Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")


    'Descripcion
    IF RS("ESTADO")=3 THEN
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    ELSE
    sInput=RS(1)
    END IF
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    SCOMSEP = chr(9)
								    scomK = "3" & SCOMSEP & "4"
								     scomD = "VIGENTE" & SCOMSEP & "NO&nbsp;VIGENTE"
								     if bModificar then
								     	sInput = ofv.gencomboFX(sForm,"ESTADO",RS("ESTADO"),scomK,scomD,"")
								     ELSE
								        IF  RS("ESTADO")= 3 THEN
								        sInput="VIGENTE"
								        ELSE
								        sInput="NO VIGENTE"
								        END IF
							         END IF
     Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Frecuencia&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  
  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

  'Botonera
  Agregar = ""
  'if bAgregar then
  '  sAccion = "AgregarFrecuencia.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si"
  '  sAccion = sAccion & "&tabla=Frecuencia"
  '  Agregar = ofv.BotonAgregar(sAccion)
  'end if

'  str1 = "Select DESCRI From Frecuencia Order By DESCRI "    ' orden pagina
'  str2 = "Select DESCRI From Frecuencia Order By DESCRI "    ' orden alfabetico
'  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sinput = "<a  href=FrecuenciaCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  sAccion = "FrecuenciaCons.asp?" & request.querystring
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