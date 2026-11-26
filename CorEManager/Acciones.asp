<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
 
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 11
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn2)


  sHTML = CorSetMHeader()
 sHTML = ofv.FormHeader(session("FormN"))
  'Encabezados
  Formulario = ofv.Encabezados("accion","tipo","Icono","Propiedades","ESTADO")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From Tipoaccalt Order By descri"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Tipoaccalt&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Descripcion
     IF RS("ESTADO")=3 THEN
       sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"150","dt","")
     else
       sInput=rs(1)
     end if
    Celda = Celda & ofv.GenCelda("","vt","","40%","","","",sInput,"si")

    Strsql1 = "Select archivo, descri From Modo where "
    Strsql1 = Strsql1 & " tipomodo = '" & RS(2) & "' "
  ''  RESPONSE.WRITE STRSQL1
    set rs1 = ofv.crearconsultaEx(StrSql1,cn,1,parametros)
    if Not rs1.EOF then
       XSMODOd = RS1(1)
       XSMODOI = ofv.GenImage("","","","","Iconos/" & rs1(0),"","")
    ELSE
       XSMODOd = "S/D"
       XSMODOI = "&nbsp;"
    END IF


    'Modo
    sInput = XSMODOd
    Celda = Celda & ofv.GenCelda("","vt","","40%","","","",sInput,"si")

    'Icono
    sImagen = XSMODOI
    Celda = Celda & ofv.GenCelda("","ft","center","10%","","","",sImagen,"si")

    IF RS("ESTADO")=3 THEN
     'sImagenP = "&#0232;"
     'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
     sinput = "<a href='AccionesProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
     sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
     sinput = sinput  & "' title='Propiedades de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
    else
     sinput="&nbsp;"
    end if


    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

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
      sAccion = sAccion & "&tabla=Tipoaccalt&volver=" & vuelta
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
    sAccion = "AgregarAcciones.asp?Regisvuelta=" & CantRegMoverx & "&tabla=Tipoaccalt&Indice=NO"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  str1 = "Select descri From Tipoaccalt Order By descri "    ' orden pagina
  str2 = "Select descri From Tipoaccalt Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  'response.write request.querystring


  sAccion = "AccionesCons.asp?" & request.querystring

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