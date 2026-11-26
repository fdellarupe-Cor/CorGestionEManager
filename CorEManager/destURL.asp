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

  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
  opc = request.querystring("opc")
  sControl = request.querystring("control")

  set cn = ofv.conectar(ofv.strconn2)

  if opc = "ins" then

	sql = "insert into destinosurl(destino,url,mail1,mail2,estado) values ('','','','','4')"
	call ofv.crearconsultaEx(sql,cn,1,parametros)

	opc = ""

  end if


	Formulario = ofv.Encabezados("Id","Destino","Url","Mail 1","Mail 2","Estado")

	sql = "select * from destinosurl order by id asc"
	set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

	do while not rs.eof

	cId = rs(0)

	sForm = "sForm" & cId

	sEsta = seEstaUsando(cId)

	if bSuprimir and sEsta = "NO" then
	   sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(cId," ","_")
	   sHref = sHref & "&Tabla=destinosurl&volver=" & vuelta & "&Eliminar=NSNC"
	   Celda = ofv.BotonEliminar(rs(1),sHref)
	else
	   Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

	Celda = Celda & ofv.gencelda("","a","right","","","","",cId,"si")

	'sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"70","dt","")
	sInput = "<input type='text' id='" & rs(1).name & "' name='" & rs(1).name & "' value='" & rs(1) & "' size=40 class=dt onblur=""document." & sForm & ".submit();"">"
	Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

	'sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"80","dt","")
	sInput = "<input type='text' id='" & rs(2).name & "' name='" & rs(2).name & "' value='" & rs(2) & "' size=45 class=dt onblur=""document." & sForm & ".submit();"">"
	Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

	'sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"100","dt","")
	sInput = "<input type='text' id='" & rs(3).name & "' name='" & rs(3).name & "' value='" & rs(3) & "' size=40 class=dt onblur=""document." & sForm & ".submit();"">"
	Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

	'sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"100","dt","")
	sInput = "<input type='text' id='" & rs(4).name & "' name='" & rs(4).name & "' value='" & rs(4) & "' size=40 class=dt onblur=""document." & sForm & ".submit();"">"
	Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

	scomK = "3" & chr(9) & "4"
	scomD = "Vigente" & chr(9) & "No&nbsp;vigente"
	sInput = ofv.gencomboFX(sForm,"estado",RS("estado"),scomK,scomD,"")
	Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

	Fila = ofv.genrow("","","","","","","",Celda,"si")

	usado = destinoControl(cId)

	sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
    sAccion = sAccion & "&tabla=destinosurl&usado=" & usado & "&volver=" & vuelta
	Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

	rs.movenext
	loop

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")




Agregar = ""
if bAgregar then
   sAccion = "destURL.asp?opc=ins"
   Agregar = ofv.BotonAgregar(sAccion)
end if


Buscar = ""
Celda = ""

sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)



	IF sControl = "SI" THEN

		'sHTML = sHTML & "<script>alert('Este destino esta siendo usado, no puede ser eliminado');</script>"
		sHTML = sHTML & "<script>alert('Existen Formularios en Estado *Desarrollo* o en Estado *En Aprobacion* que estan utilizando este destino. Elimine los destinos de los Formularios indicados para poder dar de baja este destino.');</script>"
	END IF

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

Function seEstaUsando(id)

xsql = "select iddestino from envioformdet2"
set ro = ofv.crearconsultaEx(xsql,cn,1,parametros)

xEsta = "NO"

do while not ro.eof

	if ro(0) = id then
		xEsta = "SI"
	end if

ro.movenext
loop

seEstaUsando = xEsta

End Function


Function destinoControl(xid)

ssql = "select a.iddestino from envioformdet2 a "
ssql = ssql & " inner join objprty b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto) "
ssql = ssql & " where a.iddestino=" & xid & " and b.estado in (1,2)"
set ry = ofv.crearconsultaEx(ssql,cn,1,parametros)

if ry.eof then
	sEsta = "NO"
else
	sEsta = "SI"
end if

destinoControl = sEsta

End Function

%>