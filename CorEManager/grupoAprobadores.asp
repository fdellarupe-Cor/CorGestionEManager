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
  OPCION = request.querystring("OPCION")


  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)


  'Encabezados
  Formulario = ofv.Encabezados("ID","DESCRIPCION","GRUPO","ESTADO")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  if OPCION = "INS" then

sql = "insert into gruposaprob (estado) values('1')"
call ofv.crearconsultaex(sql,cn,1,parametros)

  end if

  Strsql = "Select * From gruposAprob left join grupos on (grupos.id=gruposaprob.mapgrp)"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then rs.move(CantRegMoverx)

  posI = 1

  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    '*
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & rs(0)
      sHref = sHref & "&Tabla=gruposAprob&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)

      celda = replace(celda,"<IMG  width=15  height=15  src=" & chr(34) & "imagenes/botdel.jpg" & chr(34) & "  border=0 >","<font face=corflaticon size=2 color=#aa0000>4</font>")
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


	'ID
	sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","5%","","","",sInput,"si")

    'Descripcion
    sInput = rs("descripcion")
    Celda = Celda & ofv.gencelda("","cc","","30%","","","",sInput,"si")


 'Grupo
    sql = "select id, descripcion from grupos"
	set rsx1 = ofv.crearconsultaex(sql,cn,1,parametros)

    scomK   = ""
	scomD   = ""
	SCOMSEP = chr(9)

	do while not rsx1.eof
		scomK = scomK & SCOMSEP & rsx1(0)
		scomD = scomD & SCOMSEP & rsx1(1)
		rsx1.movenext
	loop

	call ofv.cerrarconn(rsx1)

	sinput = ofv.gencomboFX(sForm,rs(1).name,rs(1),scomK,scomD,"")
    Celda = Celda & ofv.gencelda("","cc","","30%","","","",sInput,"si")



	'Estado
	scomK = "0" & chr(9) & "1"
	scomD = "NO VIGENTE" & chr(9) & "VIGENTE"

    sinput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD,"")
    Celda = Celda & ofv.gencelda("","cc","","12%","","","",sInput,"si")




    Fila = ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=gruposaprob&volver=" & vuelta
    end if

    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

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
    sAccion = "grupoAprobadores.asp?OPCION=INS&Regisvuelta=" & CantRegMoverx
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select isnull(descripcion,'xx') From gruposAprob left join grupos on (gruposAprob.mapgrp=grupos.ID) Order By descripcion"    ' orden pagina
  str2 = str1
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))



  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>