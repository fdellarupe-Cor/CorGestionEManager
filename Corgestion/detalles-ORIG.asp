<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = "<html>"
html = html & "<head>"

html = html & "<title>Cor Gestion</title>"
html = html & "<meta charset='utf-8' />"
html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
html = html & "<link rel='stylesheet' href='assets/js/jquery-ui.css'>"
html = html & "<script type='text/javascript' src='assets/js/jquery.min.js'></script>"
html = html & "<script type='text/javascript' src='assets/js/jquery-ui.js'></script>"
html = html & "</head>"
html = html & "<body>"

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

id = request.querystring("id")
opc = request.querystring("opc")
sol = request.querystring("sol")
tipo = request.querystring("tipo")
idproy = request.querystring("idproy")
vue = request.querystring("vue")
cons = request.querystring("cons")
xtipo = request.querystring("xtipo")
cli = request.querystring("cli")
estado = request.querystring("estado")
aplic = request.querystring("aplic")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

titP = getName(cnn,"titulo","proyectos",id)
tit = getName(cnn,"titulo","incidentes",id)

nFecini = 0

If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

If ( opc = "acti" Or opc = "actp" Or opc = "veri" Or opc = "verp" ) Then

	html = html & "<div>"

	html = html & "<header id='header' style='height:11%'>"

	html = html & "<div>"

html = html & "<table style='width:100%;'><tr>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/logocorP4.gif' width='10%' >"
html = html & "<font style='margin-left:100px;font-size:30pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"
html = html & "</a></td>"

html = html & "</tr></table>"

	html = html & "</div>"

	html = html & "</header>"

	html = html & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

	html = html & "<table style='height:100%;width:100%;'>"

	html = html & "<tr>"

	If opc = "veri" Then

		html = html & "<td><font style='margin-left:5%;font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>'" & LEFT(tit,40) & "'</b></font></td>"

	ElseIf opc = "acti" Then

		html = html & "<td><font style='margin-left:5%;font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>'" & LEFT(tit,40) & "'</b></font></td>"
		cliente = request.querystring("cli")
		form5 = "<form id='form5' style='height:70%;' method='post' action='incidentes.asp?opc=listar&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		boton5 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"
		html = html & "<td>" & boton5 & "</td>"

	ElseIf opc = "actp" Then

		If tipo = "cons" Then

			html = html & "<td><font style='margin-left:5%;font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>'" & LEFT(titP,40) & "'</b></font></td>"

		Else

			sol = request.querystring("sol")
			html = html & "<td><font style='margin-left:5%;font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>'" & LEFT(titP,40) & "'</b></font></td>"

		End If
		cliente = request.querystring("cli")
		formz1 = "<form id='formz1' style='height:70%;' method='post' action='proyectos.asp?opc=listar&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		botonz1 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formz1').submit();"">"
		html = html & "<td>" & botonz1 & "</td>"

	ElseIf opc = "verp" Then

		If tipo = "cons" Then

			html = html & "<td><font style='margin-left:5%;font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>'" & LEFT(titP,40) & "'</b></font></td>"

		Else

			html = html & "<td><font style='margin-left:5%;font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>'" & LEFT(titP,40) & "'</b></font></td>"

		End If

	End If

	If opc = "veri" Then

		If tipo = "cons" Then

			cli = request.querystring("cli")
			est = request.querystring("est")
			aplic = request.querystring("apli")
			fec = request.querystring("fec")
			descri = request.querystring("desc")
			form7 = "<form id='form7' style='height:70%;' method='post' action='incidentes.asp?opc=consi&cli=" & cli & "&est=" & est & "&apli=" & aplic & "&fec=" & fec & "&desc=" & descri & "&xmenu=12&xtmenu=INICIO'>"
			boton7 = "<input class='bot' style='height:70%;font-size:13pt;' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();""></form>"
			html = html & "<td>" & form7 & boton7 & "</td>"

		elseif tipo = "relac" then

			form7 = "<form id='form7' style='height:70%;' method='post' action='detalles.asp?opc=verp&id=" & idproy & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
			boton7 = "<input class='bot' style='height:70%;font-size:13pt;' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();""></form>"
			html = html & "<td>" & form7 & boton7 & "</td>"

		elseif tipo = "incxp" then

			form7 = "<form id='form7' style='height:70%;' method='post' action='proyRelac.asp?" & request.querystring & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
			boton7 = "<input class='bot' style='height:70%;font-size:13pt;' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();""></form>"
			html = html & "<td>" & form7 & boton7 & "</td>"

		Else
			cli = request.querystring("cli")
			tit = request.querystring("tit")

			ACTpuede = puedeAct(cnn,id)

			if ACTpuede = "SI" then
				xform7 = "<form id='xform7' style='height:70%;' method='post' action='detalles.asp?opc=acti&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
				if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
					xboton7 = "<input class='bot' type='button' value='ACTUALIZAR INCIDENTE' onclick=""document.getElementById('xform7').submit();"">"
				else
					xboton7 = "<input class='botsp' type='button' value='ACTUALIZAR INCIDENTE' onclick=""alert('Sin permiso para modificar.');"">"
				end if
				html = html & "<td>" & xboton7 & "</td>"
			end if

			form7 = "<form id='form7' style='height:70%;' method='post' action='incidentes.asp?opc=listar&cli=" & cli & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
			boton7 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();"">"
			html = html & "<td >" & boton7 & "</td>"

		End If

	ElseIf opc = "verp" Then

		If tipo = "cons" Then

			cli = request.querystring("cli")
			est = request.querystring("est")
			apli = request.querystring("apli")
			fecha = request.querystring("fec")
			descri = request.querystring("desc")
			form8 = "<form id='form8' style='height:70%;' method='post' action='proyectos.asp?sel=1&cli=" & cli & "&est=" & est & "&apli=" & apli & "&fec=" & fecha & "&desc=" & descri & "&opc=consi&xmenu=12&xtmenu=INICIO'>"
			boton8 = "<input class='bot' style='height:70%;font-size:13pt;' type='button' value='VOLVER' onclick=""document.getElementById('form8').submit();""></form>"
			html = html & "<td>" & form8 & boton8 & "</td>"

		Else
			cli = request.querystring("cli")
			tit = request.querystring("tit")

			xform7 = "<form id='xform7' style='height:70%;' method='post' action='detalles.asp?opc=actp&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
			if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
				xboton7 = "<input class='bot' type='button' value='ACTUALIZAR PROYECTO' onclick=""document.getElementById('xform7').submit();"">"
			else
				xboton7 = "<input class='botsp' type='button' value='ACTUALIZAR PROYECTO' onclick=""alert('Sin permiso para modificar.');"">"
			end if
			html = html & "<td>" & xboton7 & "</td>"

			form8 = "<form id='form8' style='height:70%;' method='post' action='proyectos.asp?cli=" & cli & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
			boton8 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('form8').submit();"">"
			html = html & "<td>" & boton8 & "</td>"

		End If


	End If

	html = html & "</tr>"

	html = html & "</table>"

	html = html & "</div>"

	If opc = "veri" Then

		zzz = ObjAcc(id,"I",cnn,"D")
		XXX = SYSLOGON(cnn,"CONSULTA_INC_" & id)

		sql = " select a.id,b.descripcion,c.nombre + ' ' + c.apellido,d.descripcion,a.detalles,e.descripcion,f.nombre + ' ' + f.apellido, "
		sql = sql & " g.nombre + ' ' + g.apellido,a.req_presup,h.descripcion,i.descripcion,a.fecha_ingreso,j.descripcion,k.descripcion,a.id_proyecto,c.id, "
		sql = sql & " a.ok_desarrollo,a.ok_funcional,a.ok_cliente,a.ruta,a.creacion,a.avances,a.titulo,l.nombre,a.id_apl,m.solic_nro,m.titulo "
		sql = sql & " from incidentes a "
		sql = sql & " inner join clientes b on (a.id_cliente=b.id) "
		sql = sql & " inner join contactos c on (a.contacto=c.id) "
		sql = sql & " inner join aplicativos d on (a.id_aplic=d.id) "
		sql = sql & " inner join prioridades e on (a.prioridad=e.id) "
		sql = sql & " inner join cor_responsables f on (a.respon_gestion=f.id and f.funcion='FUNC') "
		sql = sql & " left join cor_responsables g on (a.respon_tec=g.id and g.funcion='TEC') "
		sql = sql & " inner join estados h on (a.estado_desa=h.id) "
		sql = sql & " inner join estados i on (a.estado_func=i.id) "
		sql = sql & " inner join browsers j on (a.browser=j.id) "
		sql = sql & " inner join via_pedido k on (a.via_pedido=k.id) "
		sql = sql & " left join apl l on (a.id_apl=l.id) "
		sql = sql & " left join proyectos m on (a.id_proyecto = m.id) "
		sql = sql & " where a.id=" & id

		Set rs = Cnn.Execute(sql)

		If Not rs.eof Then

		html = html & "<div style='height:80%;width:100%;float:left;'>"

		html = html & "<div id='tabs' style='border-style: none;'>"

		html = html & "<ul>"
		html = html & "<li><a href='#fragment-1'>DATOS</a></li>"
		html = html & "<li><a href='#fragment-2'>DETALLE</a></li>"
		html = html & "<li><a href='#fragment-3'>AVANCE</a></li>"
		html = html & "<li><a href='#fragment-4'>APLs RELACIONADOS</a></li>"
		html = html & "</ul>"

		'DATOS
		html = html & "<div id='fragment-1'>"

		html = html & "<table class='tablDet' cellspacing=0 cellpadding=3 align=center>"

		html = html & "<tr>"
		html = html & "<td class='fil'><b>NRO</b></td><td class='fil'>" & rs(0) & "</td>"
		html = html & "<td class='fil'><b>TITULO</b></td><td class='fil'>" & rs(22) & "</td>"
		html = html & "</tr>"


		html = html & "<tr>"
		html = html & "<td class='fil'><b>CLIENTE</b></td><td class='fil'>" & rs(1) & "</td>"
		html = html & "<td class='fil'><b>CONTACTO</b></td>"
		html = html & "<td class='fil'><a class='red' href='detalleCont.asp?idcon=" & rs(15) & "&id=" & rs(0) & "&opc=veri&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'><font style='color:#FF0000;'>" & rs(2) & "</font></a></td>"
		html = html & "</tr>"

		sRelPro = ""
		if rs(14) <> 0 then sRelPro = "ID: " & rs(14)

		html = html & "<tr>"
		html = html & "<td class='fil'><b>APLICACION</b></td><td class='fil'>" & rs(3) & "</td>"
		html = html & "<td class='fil'><b>RELACIONADO A PROY</b></td><td class='fil'><a class='red' href='detalles.asp?id=" & rs(14) & "&opc=verp&sol=" & rs(25) & "&xmenu=10&xtmenu=INICIO'><font style='color:#FF0000;'>" & sRelPro & "</font></a></td>"
		html = html & "</tr>"



		html = html & "<tr>"
		html = html & "<td class='fil'><b>PRIORIDAD</b></td><td class='fil'>" & rs(5) & "</td>"
		html = html & "<td class='fil'><b>RESPONSABLE GESTION</b></td><td class='fil'>" & rs(6) & "</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		rp = ""
		if rs(7) & "" = "" then
		rp = "---"
		else
		rp = rs(7)
		end if
		html = html & "<td class='fil'><b>RESPONSABLE TECNICO</b></td><td class='fil'>" & rp & "</td>"
		html = html & "<td class='fil'><b>REQUIERE PRESUPUESTO</b></td><td class='fil'>" & rs(8) & "</td>"
		html = html & "</tr>"


		html = html & "<tr>"
		html = html & "<td class='fil'><b>ESTADO DESARROLLO</b></td><td class='fil'>" & rs(9) & "</td>"
		html = html & "<td class='fil'><b>ESTADO FUNCIONAL</b></td><td class='fil'>" & rs(10) & "</td>"
		html = html & "</tr>"

		sFec = Fechador(rs(11) + nFecini)

		html = html & "<tr>"
		html = html & "<td class='fil'><b>FECHA CREACION</b></td><td class='fil'>" & sFec & "</td>"
		html = html & "<td class='fil'><b>BROWSER</b></td><td class='fil'>" & rs(12) & "</td>"
		html = html & "</tr>"

		od = ""
		if rs(16) & "" = "0" then
			od = "---"
		else
			od = Fechador(rs(16) + nFecini)
		end if

		html = html & "<tr>"
		html = html & "<td class='fil'><b>VIA PEDIDO</b></td><td class='fil'>" & rs(13) & "</td>"
		html = html & "<td class='fil'><b>OK DESARROLLO</b></td><td class='fil'>" & od & "</td>"
		html = html & "</tr>"

		of = ""
		if rs(17) & "" = "0" then
		of = "---"
		else
		of = Fechador(rs(17) + nFecini)
		end if

		html = html & "<tr>"
		html = html & "<td class='fil'><b>OK FUNCIONAL</b></td><td class='fil'>" & of & "</td>"
		html = html & "<td class='fil'><b>OK CLIENTE</b></td><td class='fil'>" & rs(18) & "</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'><b>LINK DOCUMENTO</b></td><td class='fil'>"
		html = html & "<a href='" & rs("ruta") & "' target='_new'><font style='color:#FF0000;'>" & Left(rs("ruta"),30) & "</font></a>"
		html = html & "</td>"
		html = html & "<td class='fil'>&nbsp;</td>"
		html = html & "<td class='fil'>&nbsp;</td>"
		html = html & "</tr>"

		html = html & "</table>"

		html = html & "</div>"

		'DETALLES
		html = html & "<div id='fragment-2'>"

		html = html & "<table class='tabldet' style='border-style: none;' cellspacing=0 align=left>"

		html = html & "<tr>"
		html = html & "<td class='fil' style='font-size:11pt;'>" & Replace(rs(4) & "",Chr(13),"<br>") & "</td>"
		html = html & "</tr>"

		html = html & "</table>"

		html = html & "</div>"

		'AVANCES
		html = html & "<div id='fragment-3'>"

		html = html & "<table class='tablDet' style='border-style: none;' cellspacing=0 align=left>"

		html = html & "<tr>"
		html = html & "<td class='fil' style='font-size:11pt;'>" & Replace(rs(21) & "",Chr(13),"<br>") & "</td>"
		html = html & "</tr>"

		html = html & "</table>"

		html = html & "</div>"

		'APLs RELACIONADOS
		html = html & "<div id='fragment-4'>"

		html = html & "<table style='border:solid 1px #aaaaaa;' cellspacing=0 cellpadding=3 align=center>"

		sql = "select * from apl_relac where id_relac=" & id & " and tipo_relac='I'"
		Set rz = Cnn.Execute(sql)

		html = html & "<tr>"
		html = html & "<td><font style='font-size:20pt;color:#000000;'>APLs</font></td>"
		html = html & "</tr>"

		if not rz.eof then

 		do while not rz.eof

		html = html & "<tr>"
		ref = "<a class='red2' href='detalles2.asp?opc=ver&id=" & rz(1) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & rz(1) & "</a>"
		html = html & "<td align=center>" & ref & "</td>"
		html = html & "</tr>"

		rz.movenext
		loop

		else

			html = html & "<tr>"
			html = html & "<td>-----</td>"
			html = html & "</tr>"

		end if

		html = html & "</table>"

		html = html & "</div>"

		html = html & "</div>"

		End If

		rs.close

	ElseIf opc = "verp" Then

		zzz = ObjAcc(id,"P",cnn,"D")
		XXX = SYSLOGON(cnn,"CONSULTA_PROY_" & id)

		sql = "select a.id,a.solic_nro,a.titulo,a.descripcion,b.descripcion,c.nombre + ' ' + c.apellido,d.descripcion,f.nombre + ' ' + f.apellido, "
		sql = sql & " g.nombre + ' ' + g.apellido,a.fecha_ingreso,a.fecha_entrega,h.descripcion,i.descripcion,e.descripcion,a.tiempo_relevamiento,a.tiempo_diseno, "
		sql = sql & " a.tiempo_desarr,a.tiempo_test, "
		sql = sql & " j.descripcion,k.descripcion,a.req_presup,SUM(a.tiempo_diseno + a.tiempo_desarr + a.tiempo_relevamiento + a.tiempo_test),c.id, "
		sql = sql & " a.ok_desarrollo,a.ok_funcional,a.ok_cliente,a.ruta,a.avances,l.nombre,a.id_apl "
		sql = sql & " from proyectos a "
		sql = sql & " inner join clientes b on (a.id_cliente=b.id) "
		sql = sql & " inner join contactos c on (a.contacto=c.id) "
		sql = sql & " inner join aplicativos d on (a.id_aplic = d.id) "
		sql = sql & " inner join prioridades e on (a.prioridad=e.id) "
		sql = sql & " inner join cor_responsables f on (a.respon_gestion=f.id and f.funcion='FUNC') "
		sql = sql & " inner join cor_responsables g on (a.respon_tec=g.id and g.funcion='TEC') "
		sql = sql & " inner join estados h on (a.estado_desa=h.id) "
		sql = sql & " inner join estados i on (a.estado_func=i.id) "
		sql = sql & " inner join browsers j on (a.browser=j.id) "
		sql = sql & " inner join via_pedido k on (a.via_pedido=k.id) "
		sql = sql & " left join apl l on (a.id_apl = l.id) "
		sql = sql & " where a.id=" & id
		sql = sql & " group by a.id,a.solic_nro,a.titulo,a.descripcion,b.descripcion,c.nombre + ' ' + c.apellido,d.descripcion,"
		sql = sql & " f.nombre + ' ' + f.apellido,g.nombre + ' ' + g.apellido, a.fecha_ingreso,a.fecha_entrega,h.descripcion,"
		sql = sql & " i.descripcion,e.descripcion,a.tiempo_relevamiento,a.tiempo_diseno,a.tiempo_desarr,a.tiempo_test, "
		sql = sql & " j.descripcion,k.descripcion,a.req_presup,a.id_cliente,a.contacto,a.id_aplic,a.respon_gestion,a.respon_tec,"
		sql = sql & " a.estado_func,a.estado_desa,a.prioridad, a.browser,a.via_pedido,a.req_presup,c.id,a.ok_desarrollo,a.ok_funcional, "
		sql = sql & " a.ok_cliente,a.ruta,a.avances,l.nombre,a.id_apl"

		Set rs = Cnn.Execute(sql)

		If Not rs.eof Then

			html = html & "<div style='height:80%;width:100%;float:left;'>"

			html = html & "<div id='tabs' style='border-style: none;' align=center>"

			html = html & "<ul>"
			html = html & "<li><a href='#fragment-1'><span>DATOS</span></a></li>"
			html = html & "<li><a href='#fragment-2'><span>DETALLE</span></a></li>"
			html = html & "<li><a href='#fragment-3'><span>AVANCE</span></a></li>"
			html = html & "<li><a href='#fragment-4'><span>INCIDENTES Y APLs RELACIONADOS</span></a></li>"
			html = html & "</ul>"

			'DATOS
			html = html & "<div id='fragment-1'>"

			html = html & "<table class='tablDet' cellspacing=0 cellpadding=3 align=center>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>ID</b></td><td class='fil'>" & rs(0) & "</td>"
			html = html & "<td class='fil'><b>SOLICITUD NUMERO</b></td><td class='fil'>" & rs(1) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>TITULO PROYECTO</b></td><td class='fil'>" & rs(2) & "</td>"
			html = html & "<td class='fil'><b>CLIENTE</b></td><td class='fil'>" & rs(4) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>CONTACTO</b></td>"
			html = html & "<td class='fil'><a class='red' href='detalleCont.asp?id=" & id & "&opc=" & opc & "&sol=" & sol & "&idcon=" & rs(22) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'><font style='color:#FF0000;'>" & rs(5) & "</font></a></td>"
			html = html & "<td class='fil'><b>APLICACION</b></td><td class='fil'>" & rs(6) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>RESPON GESTION</b></td><td class='fil'>" & rs(7) & "</td>"
			html = html & "<td class='fil'><b>RESPON TEC</b></td><td class='fil'>" & rs(8) & "</td>"
			html = html & "</tr>"

			sFec = ""
			if rs(9) <> 0 then
				sFec = Fechador(rs(9) + nFecini)
			end if

			sFecEn = ""
			if rs(10) <> 0 then
				sFecEn = Fechador(rs(10) + nFecini)
			end if

			html = html & "<tr>"
			html = html & "<td class='fil'><b>FECHA CREACION</b></td><td class='fil'>" & sFec & "</td>"
			html = html & "<td class='fil'><b>FECHA ENTREGA</b></td><td class='fil'>" & sFecEn & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>ESTADO FUNCIONAL</b></td><td class='fil'>" & rs(12) & "</td>"
			html = html & "<td class='fil'><b>ESTADO DESARROLLO</b></td><td class='fil'>" & rs(11) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>PRIORIDAD</b></td><td class='fil'>" & rs(13) & "</td>"
			html = html & "<td class='fil'><b>RELEV Y DEFINICION (hs)</b></td><td class='fil'>" & rs(14) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>DISENO (hs)</b></td><td class='fil'>" & rs(15) & "</td>"
			html = html & "<td class='fil'><b>DESARROLLO (hs)</b></td><td class='fil'>" & rs(16) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>TESTEO (hs)</b></td><td class='fil'>" & rs(17) & "</td>"
			html = html & "<td class='fil'><b>TOTAL (hs)</b></td><td class='fil'>" & rs(21) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>BROWSER</b></td><td class='fil'>" & rs(18) & "</td>"
			html = html & "<td class='fil'><b>VIA PEDIDO</b></td><td class='fil'>" & rs(19) & "</td>"
			html = html & "</tr>"

			od = ""
			if rs(23) & "" = "0" then
				od = "---"
			else
				od = Fechador(rs(23) + nFecini)
			end if
			html = html & "<tr>"
			html = html & "<td class='fil'><b>REQ PRESUPUESTO</b></td><td class='fil'>" & rs(20) & "</td>"
			html = html & "<td class='fil'><b>OK DESARROLLO</b></td><td class='fil'>" & od & "</td>"
			html = html & "</tr>"

			of = ""
			if rs(24) & "" = "0" then
				of = "---"
			else
				of = Fechador(rs(24) + nFecini)
			end if
			html = html & "<tr>"
			html = html & "<td class='fil'><b>OK FUNCIONAL</b></td><td class='fil'>" & of & "</td>"
			html = html & "<td class='fil'><b>OK CLIENTE</b></td><td class='fil'>" & rs(25) & "</td>"
			html = html & "</tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'><b>LINK DOCUMENTO</b></td><td class='fil'>"
			html = html & "<a href='" & rs("ruta") & "' target='_new'><font style='color:#FF0000;'>" & Left(rs("ruta"),30) & "</font></a>"
			html = html & "</td>"
			html = html & "</tr>"

			html = html & "</table>"

			html = html & "</div>"

			'DETALLES
			html = html & "<div id='fragment-2'>"

			html = html & "<table cellspacing=0 align=left>"

			html = html & "<tr>"
			html = html & "<td class='fil' style='font-size:11pt;'>" & Replace(rs(3) & "",Chr(13),"<br>") & "</td>"
			html = html & "</tr>"

			html = html & "</table>"

			html = html & "</div>"

			'AVANCES
			html = html & "<div id='fragment-3'>"

			html = html & "<table cellspacing=0 align=left>"

			html = html & "<tr>"
			html = html & "<td class='fil' style='font-size:11pt;'>" & Replace(rs(27) & "",Chr(13),"<br>") & "</td>"
			html = html & "</tr>"

			html = html & "</table>"

			html = html & "</div>"



			html = html & "<div id='fragment-4' style='width:300px;' align=center>"

			html = html & "<table style='border:solid 1px #aaaaaa;float:left;' cellspacing=0 cellpadding=3 align=center>"

			html = html & "<tr>"
			html = html & "<td><font style='font-size:20pt;color:#000000;'>APLs</font></td>"
			html = html & "</tr>"

			sql = "select * from apl_relac where id_relac=" & id & " and tipo_relac='P'"
			Set rz = Cnn.Execute(sql)

			if not rz.eof then

			do while not rz.eof

				html = html & "<tr>"
				ref = "<a class='red2' href='detalles2.asp?opc=ver&id=" & rz(1) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & rz(1) & "</a>"
				html = html & "<td align=center>" & ref & "</td>"
				html = html & "</tr>"

			rz.movenext
			loop

			else

				html = html & "<tr>"
				html = html & "<td>-----</td>"
				html = html & "</tr>"

			end if

			html = html & "</table>"

			html = html & "<table style='border:solid 1px #aaaaaa;float:right;' cellspacing=0 cellpadding=3 align=center>"

			html = html & "<tr>"
			html = html & "<td><font style='font-size:20pt;color:#000000;'>INCIDENTES</font></td>"
			html = html & "</tr>"

			sql = "select * from incidentes where id_proyecto=" & id
			Set rz = Cnn.Execute(sql)

			if not rz.eof then

			do while not rz.eof

				html = html & "<tr>"
				ref = "<a class='red2' href='detalles.asp?opc=veri&id=" & rz(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & rz(0) & "</a>"
				html = html & "<td align=center>" & ref & "</td>"
				html = html & "</tr>"

			rz.movenext
			loop

			else

				html = html & "<tr>"
				html = html & "<td>-----</td>"
				html = html & "</tr>"

			end if

			html = html & "</table>"

			html = html & "</div>"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


			html = html & "</div>"

			html = html & "<div style='height:100%;width:50%;float:right'><br>"






		End If

		html = html & "</table>"

	ElseIf opc = "acti" Then

		zzz = ObjAcc(id,"I",cnn,"A")
		XXX = SYSLOGON(cnn,"ACTUALIZA_INC_" & id)

		html = html & "<div style='height:80%;width:50%;float:left;'><br>"

		html = html & "<form id='sForm' method='post' action='actualizar.asp?tabla=incidentes&id=" & id & "&opt=actinc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=acti' >"

		html = html & "<table class='tablDet' style='width:70%;' cellspacing=0 align=center>"

		sql = " select a.id,b.id,b.descripcion,c.id,c.nombre + ' ' + c.apellido,d.id,d.descripcion,e.id,a.detalles,e.descripcion, "
		sql = sql & " f.id,f.nombre + ' ' + f.apellido,ISNULL(g.id,'0'),ISNULL(g.nombre + ' ' + g.apellido,'0'),a.req_presup,h.id,h.descripcion,i.id,i.descripcion,a.fecha_ingreso, "
		sql = sql & " j.id,j.descripcion,k.id,k.descripcion,a.id_proyecto,a.ok_cliente,a.ruta,a.avances,a.titulo,a.id_apl,m.solic_nro,m.titulo "
		sql = sql & " from incidentes a "
		sql = sql & " inner join clientes b on (a.id_cliente=b.id) "
		sql = sql & " left join contactos c on (a.contacto=c.id) "
		sql = sql & " inner join aplicativos d on (a.id_aplic=d.id) "
		sql = sql & " inner join prioridades e on (a.prioridad=e.id) "
		sql = sql & " inner join cor_responsables f on (a.respon_gestion=f.id and f.funcion='FUNC') "
		sql = sql & " left join cor_responsables g on (a.respon_tec=g.id and g.funcion='TEC') "
		sql = sql & " inner join estados h on (a.estado_desa=h.id) "
		sql = sql & " inner join estados i on (a.estado_func=i.id) "
		sql = sql & " inner join browsers j on (a.browser=j.id) "
		sql = sql & " inner join via_pedido k on (a.via_pedido=k.id) "
		sql = sql & " left join proyectos m on (a.id_proyecto=m.id) "
		sql = sql & " where a.id=" & id

		Set rx = Cnn.Execute(sql)

		html = html & "<tr>"
		html = html & "<td class='fil' colspan=2 style='font-size:12pt;'><b><u>DATOS</u></b></td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>NRO</td><td id='nid' class='fil'>" & id & "</td>"
		html = html & "</tr>"

		if rx(10) = session("USUARIO") and session("FUNC") = "FUNC" then

			html = html & "<tr>"
			html = html & "<td class='fil'>CLIENTE</td><td class='fil'>"
			html = html & armaCombo("select * from clientes",Cnn,rx(1),0,1,"id_cliente")
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>TITULO</td><td class='fil'>"
			html = html & "<input class='inpu' id='titulo' name='titulo' type='text' style='width:90%;' value='" & rx(28) & "'>"
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>CONTACTO</td><td class='fil'>"
			html = html & armaCombo("select id,nombre + ' ' + apellido from contactos",Cnn,rx(3),0,1,"contacto")
			html = html &"</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>APLICACION</td><td class='fil'>"
			html = html & armaCombo("select id,descripcion from aplicativos",Cnn,rx(5),0,1,"id_aplic")
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>RELACIONADO A PROYECTO</td><td class='fil'>"
			html = html & armaCombo("select id,solic_nro + ' - ' + titulo from proyectos where id_cliente=" & rx(1),Cnn,rx(24),0,1,"id_proyecto")
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>PRIORIDAD</td><td class='fil'>"
			html = html & armaCombo("select id,descripcion from prioridades",Cnn,rx(7),0,1,"prioridad")
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>RESPONSABLE GESTION</td><td class='fil'>"
			html = html & armaCombo("select id,nombre + ' ' + apellido from cor_responsables where funcion='FUNC'",Cnn,rx(10),0,1,"respon_gestion")
			html = html & "</td></tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>CLIENTE</td><td class='fil'>"
			html = html & rx(2)
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>CONTACTO</td><td class='fil'>"
			html = html & rx(4)
			html = html &"</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>APLICACION</td><td class='fil'>"
			html = html & rx(6)
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>RELAC A PROY</td><td class='fil'>"
			html = html & rx(30) & "-" & rx(31)
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>PRIORIDAD</td><td class='fil'>"
			html = html & rx(9)
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>RESPONSABLE GESTION</td><td class='fil'>"
			html = html & rx(11)
			html = html & "</td></tr>"

		end if

		if (rx(12) = session("USUARIO") and session("FUNC") = "TEC") then

			html = html & "<tr>"
			html = html & "<td class='fil'>RESPONSABLE TECNICO</td><td class='fil'>"
			html = html & armaCombo("select id,nombre + ' ' + apellido from cor_responsables where funcion='TEC'",Cnn,rx(12),0,1,"respon_tec")
			html = html & "</td></tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>RESPONSABLE TECNICO</td><td class='fil'>"
			rp = ""
			if rx(13) & "" = "0" then
				rp = "---"
			else
				rp = rx(13)
			end if
			html = html & rp
			html = html & "</td></tr>"

		end if

		if rx(10) = session("USUARIO") and session("FUNC") = "FUNC" then

			html = html & "<tr>"
			html = html & "<td class='fil'>REQUIERE PRESUPUESTO</td><td class='fil'>"
			html = html & "<select class='selec' style='width:90%;' id='req_presup' name='req_presup'>"

			If rx(14) = "SI" Then

				html = html & "<option value='SI' selected>SI</option>"
				html = html & "<option value='NO'>NO</option>"

			Else

				html = html & "<option value='SI'>SI</option>"
				html = html & "<option value='NO' selected>NO</option>"

			End If

			html = html & "</select></td></tr>"

			sFec = Fechador(rx(19) + nFecini)

			html = html & "<td class='fil'>FECHA INGRESO</td><td class='fil'>"

			html = html & sFec

			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>BROWSER</td><td class='fil'>"
			html = html & armaCombo("select id,descripcion from browsers",Cnn,rx(20),0,1,"browser")
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>VIA PEDIDO</td><td class='fil'>"
			html = html & armaCombo("select id,descripcion from via_pedido",Cnn,rx(22),0,1,"via_pedido")
			html = html & "</td></tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>REQUIERE PRESUPUESTO</td><td class='fil'>"
			html = html & rx(14) & "</td></tr>"

			sFec = Fechador(rx(19) + nFecini)

			html = html & "<td class='fil'>FECHA INGRESO</td><td class='fil'>"
			html = html & sFec
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>BROWSER</td><td class='fil'>"
			html = html & rx(21)
			html = html & "</td></tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>VIA PEDIDO</td><td class='fil'>"
			html = html & rx(23)
			html = html & "</td></tr>"

		end if

		if rx(12) = session("USUARIO") and session("FUNC") = "TEC" then

			html = html & "<tr>"
			html = html & "<td class='fil'>ESTADO DESARROLLO</td><td class='fil'>"
			html = html & armaCombo("select id,descripcion from estados",Cnn,rx(15),0,1,"estado_desa")
			html = html & "</td></tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>ESTADO DESARROLLO</td><td class='fil'>"
			html = html & rx(16)
			html = html & "</td></tr>"

		end if

	if rx(10) = session("USUARIO") and session("FUNC") = "FUNC" then

		if rx(16) & "" = "OK" OR rx(16) & "" = "BAJA" or rx(16) & "" = "---" then

			html = html & "<tr>"
			html = html & "<td class='fil'>ESTADO FUNCIONAL</td><td class='fil'>"
			html = html & armaCombo("select id,descripcion from estados",Cnn,rx(17),0,1,"estado_func")
			html = html & "</td></tr>"
			html = html & "<tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>ESTADO FUNCIONAL</td><td class='fil'>"
			html = html & rx(18)
			html = html & "</td></tr>"
			html = html & "<tr>"

		end if

		if rx(16) & "" = "OK" and rx(18) & "" = "OK" then

			html = html & "<tr>"
			html = html & "<td class='fil'>FECHA OK CLIENTE</td><td class='fil'>"
			html = html & "<input id='ok_cliente' name='ok_cliente' type='text' class='inpu' style='width:90%;' value='" & rx(25) & "'>"
			html = html & "</td></tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>FECHA OK CLIENTE</td><td class='fil'>"
			html = html & rx(25)
			html = html & "</td></tr>"

		end if

	else

			html = html & "<tr>"
			html = html & "<td class='fil'>ESTADO FUNCIONAL</td><td class='fil'>"
			html = html & rx(18)
			html = html & "</td></tr>"
			html = html & "<tr>"

			html = html & "<tr>"
			html = html & "<td class='fil'>FECHA OK CLIENTE</td><td class='fil'>"
			html = html & rx(25)
			html = html & "</td></tr>"

	end if

		html = html & "<tr>"
		html = html & "<td class='fil'>LINK DOCUMENTO</td><td class='fil'>"
		test = "<a href='" & rx("ruta") & "'><font style='color:#FF0000;'>Test</font></a>"
		html = html & "<input class='inpu' style='width:90%;' id='ruta' name='ruta' size=50 type='text' value='" & rx("ruta") & "'>" & test
		html = html & "</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td colspan=2>&nbsp;</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td id='sAgre' class='fil' colspan=2 align=center><font style='color:#FF0000;'>RELACIONAR APL</font></td>"
		html = html & "</tr>"

		html = html & "</table>"

		html = html & "<br><font style='color:#000000;font-size:12pt;margin-left:100px;'>"
		html = html & "<u>ejemplo link:</u> 'http://server-2025/registroincidente/' + nombre documento"
		html = html & "</font>"

		html = html & "</div>"

		html = html & "<div style='height:100%;width:50%;float:right;'><br>"

		html = html & "<table class='tablDet' style='width:80%;' cellspacing=0 align=left>"

		html = html & "<tr>"
		html = html & "<td class='fil' style='font-size:12pt;'><u><b>DETALLES</b></u></td>"
		html = html & "</tr>"

		if rx(10) = session("USUARIO") and session("FUNC") = "FUNC" then

			html = html & "<tr>"
			html = html & "<td class='fil'>"
			html = html & "<textarea class='txtar' id='detalles' name='detalles' type='text' style='height:200px;width:100%;'>" & rx(8) & "</textarea>"
			html = html & "</td></tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>"
			html = html & rx(8)
			html = html & "</td></tr>"

		end if

		html = html & "</table>"

		html = html & "<table class='tablDet' style='margin-top:50px;width:80%;' cellspacing=0 align=left>"

		html = html & "<tr>"
		html = html & "<td class='fil' style='font-size:12pt;'><b><u>AVANCES</u></b></td>"
		html = html & "</tr>"

		if (rx(10) = session("USUARIO") and session("FUNC") = "FUNC") or (rx(12) = session("USUARIO") and session("FUNC") = "TEC") then

			html = html & "<tr>"
			html = html & "<td class='fil' xstyle='border:solid 1px #001BFF;'>"
			html = html & "<textarea class='txtar' id='avances' name='avances' type='text' style='height:200px;width:100%;'>" & rx(27) & "</textarea>"
			html = html & "</td>"
			html = html & "</tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil' xstyle='border:solid 1px #001BFF;'>" & rx(27) & "</td>"
			html = html & "</tr>"

		end if

		html = html & "<tr>"
		html = html & "<td>&nbsp;</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td>"
		html = html & "<input class='bot' type='button' xstyle='border-radius:10px;border:solid 2px;border-color:#FF0000;height:100%;width:100%;' value='ACTUALIZAR' onclick=""actualizar('sForm');"">"
		html = html & "</td></tr>"

		html = html & "</table>"

		html = html & "</form>"

		rx.close

	ElseIf opc = "actp" Then

		zzz = ObjAcc(id,"P",cnn,"A")
		XXX = SYSLOGON(cnn,"ACTUALIZA_PROY_" & id)

		sql = "select a.id,a.solic_nro,a.titulo,a.descripcion,b.descripcion,c.nombre + ' ' + c.apellido,d.descripcion,f.nombre + ' ' + f.apellido, "
		sql= sql & " g.nombre + ' ' + g.apellido,a.fecha_ingreso,a.fecha_entrega,h.descripcion,i.descripcion,e.descripcion,a.tiempo_relevamiento,a.tiempo_diseno, "
		sql= sql & " a.tiempo_desarr,a.tiempo_test,j.descripcion,k.descripcion,a.req_presup,a.id_cliente,a.contacto,id_aplic,a.respon_gestion,a.respon_tec, "
		sql= sql & " a.estado_func,a.estado_desa,a.prioridad,a.browser,a.via_pedido,a.req_presup, "
		sql= sql & " SUM(a.tiempo_diseno + a.tiempo_desarr + a.tiempo_relevamiento + a.tiempo_test),a.ok_cliente,a.ruta,a.avances,a.id_apl "
		sql= sql & " from proyectos a "
		sql= sql & " inner join clientes b on (a.id_cliente=b.id) "
		sql= sql & " inner join contactos c on (a.contacto=c.id) "
		sql= sql & " inner join aplicativos d on (a.id_aplic=d.id) "
		sql= sql & " inner join prioridades e on (a.prioridad=e.id) "
		sql= sql & " inner join cor_responsables f on (a.respon_gestion=f.id and f.funcion='FUNC') "
		sql= sql & " inner join cor_responsables g on (a.respon_tec=g.id and g.funcion='TEC') "
		sql= sql & " inner join estados h on (a.estado_desa=h.id) "
		sql= sql & " inner join estados i on (a.estado_func=i.id) "
		sql= sql & " inner join browsers j on (a.browser=j.id) "
		sql= sql & " inner join via_pedido k on (a.via_pedido=k.id) "
		sql= sql & " where a.id=" & id
		sql= sql & " group by a.id,a.solic_nro,a.titulo,a.descripcion,b.descripcion,c.nombre + ' ' + c.apellido,d.descripcion,"
		sql= sql & " f.nombre + ' ' + f.apellido,g.nombre + ' ' + g.apellido, a.fecha_ingreso,a.fecha_entrega,h.descripcion,"
		sql= sql & " i.descripcion,e.descripcion,a.tiempo_relevamiento,a.tiempo_diseno,a.tiempo_desarr,a.tiempo_test, "
		sql= sql & " j.descripcion,k.descripcion,a.req_presup,a.id_cliente,a.contacto,id_aplic,a.respon_gestion,a.respon_tec,"
		sql= sql & " a.estado_func,a.estado_desa,a.prioridad, a.browser,a.via_pedido,a.req_presup,a.ok_cliente,a.ruta,a.avances,a.id_apl"

		Set rs = Cnn.Execute(sql)

		if not rs.eof then

		html = html & "<div style='height:80%;width:50%;float:left;'><br>"

		html = html & "<form id='sForm' method='post' action='actualizar.asp?tabla=proyectos&id=" & id & "&opt=actproy&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=actp&sol=" & rs(1) & "' >"

		html = html & "<table class='tablDet' style='width:70%;' cellspacing=0 align=center>"

		html = html & "<tr>"
		html = html & "<td class='fil' style='font-size:12pt;' colspan=2><b><u>DATOS</u></b></td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>ID</td><td id='nid' class='fil'>" & rs(0) & "</td>"
		html = html & "</tr>"

	if rs(24) = session("USUARIO") and session("FUNC") = "FUNC" then

		html = html & "<tr>"
		html = html & "<td class='fil'>SOLICITUD NUMERO</td><td class='fil'>"
		html = html & "<input class='inpu' id='solic_nro' name='solic_nro' type='text' style='width:90%;' value='" & rs(1) & "'>"
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>TITULO PROYECTO</td><td class='fil'>"
		html = html & "<input class='inpu' id='titulo' name='titulo' type='text' style='width:90%;' value='" & rs(2) & "'>"
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>CLIENTE</td><td class='fil'>"
		html = html & armaCombo("select * from clientes",Cnn,rs(21),0,1,"id_cliente")
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>CONTACTO</td><td class='fil'>"
		html = html & armaCombo("select id,nombre + ' ' + apellido from contactos",Cnn,rs(22),0,1,"contacto")
		html = html &"</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>APLICACION</td><td class='fil'>"
		html = html & armaCombo("select id,descripcion from aplicativos",Cnn,rs(23),0,1,"id_aplic")
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>RESPONSABLE GESTION</td><td class='fil'>"
		html = html & armaCombo("select id,nombre + ' ' + apellido from cor_responsables where funcion='FUNC'",Cnn,rs(24),0,1,"respon_gestion")
		html = html & "</td></tr>"

	else

		html = html & "<tr>"
		html = html & "<td class='fil'>SOLICITUD NUMERO</td><td class='fil'>"
		html = html & rs(1)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>TITULO PROYECTO</td><td class='fil'>"
		html = html & rs(2)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>CLIENTE</td><td class='fil'>"
		html = html & rs(4)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>CONTACTO</td><td class='fil'>"
		html = html & rs(5)
		html = html &"</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>APLICACION</td><td class='fil'>"
		html = html & rs(6)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>RESPONSABLE GESTION</td><td class='fil'>"
		html = html & rs(7)
		html = html & "</td></tr>"

	end if

	if rs(25) = session("USUARIO") and session("FUNC") = "TEC" then

		html = html & "<tr>"
		html = html & "<td class='fil'>RESPONSABLE TECNICO</td><td class='fil'>"
		html = html & armaCombo("select id,nombre + ' ' + apellido from cor_responsables where funcion='TEC'",Cnn,rs(25),0,1,"respon_tec")
		html = html & "</td></tr>"

	else

		html = html & "<tr>"
		html = html & "<td class='fil'>RESPONSABLE TECNICO</td><td class='fil'>"
		html = html & rs(8)
		html = html & "</td></tr>"

	end if

	if rs(24) = session("USUARIO") and session("FUNC") = "FUNC" then

		sFec = ""
		if rs(9) <> 0 then
			sFec = Fechador(rs(9) + nFecini)
		end if

		html = html & "<tr>"
		html = html & "<td class='fil'>FECHA INGRESO</td><td class='fil'>"
		html = html & sFec
		html = html & "</td></tr>"

		sFecEn = ""
		if rs(10) <> 0 then
			sFecEn = Fechador(rs(10) + nFecini)
		end if

		html = html & "<tr>"
		html = html & "<td class='fil'>FECHA ENTREGA</td><td class='fil'>"
		html = html & "<input class='inpu' id='fecha_entrega' name='fecha_entrega' type='text' style='width:90%;' value='" & sFecEn & "'>"
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>PRIORIDAD</td><td class='fil'>"
		html = html & armaCombo("select id,descripcion from prioridades",Cnn,rs(28),0,1,"prioridad")
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>RELEV Y DEFINICION (hs)</td><td class='fil'>"
		html = html & "<input class='inpu' id='tiempo_relevamiento' name='tiempo_relevamiento' type='text' style='width:90%;' value='" & rs(14) & "'>"
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>DISENO (hs)</td><td class='fil'>"
		html = html & "<input class='inpu' id='tiempo_diseno' name='tiempo_diseno' type='text' style='width:90%;' value='" & rs(15) & "'>"
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>DESARROLLO (hs)</td><td class='fil'>"
		html = html & "<input class='inpu' id='tiempo_desarr' name='tiempo_desarr' type='text' style='width:90%;' value='" & rs(16) & "'>"
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>TESTEO (hs)</td><td class='fil'>"
		html = html & "<input class='inpu' id='tiempo_test' name='tiempo_test' type='text' style='width:90%;' value='" & rs(17) & "'>"
		html = html & "</td></tr>"

	else

		sFec = Fechador(rs(9) + nFecini)

		html = html & "<tr>"
		html = html & "<td class='fil'>FECHA INGRESO</td><td class='fil'>"
		html = html & sFec
		html = html & "</td></tr>"

		sFecEn = ""
		if rs(10) <> 0 then
			sFecEn = Fechador(rs(10) + nFecini)
		end if

		html = html & "<tr>"
		html = html & "<td class='fil'>FECHA ENTREGA</td><td class='fil'>"
		html = html & sFecEn
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>PRIORIDAD</td><td class='fil'>"
		html = html & rs(13)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>RELEV Y DEFINICION (hs)</td><td class='fil'>"
		html = html & rs(14)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>DISENO (hs)</td><td class='fil'>"
		html = html & rs(15)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>DESARROLLO (hs)</td><td class='fil'>"
		html = html & rs(16)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>TESTEO (hs)</td><td class='fil'>"
		html = html & rs(17)
		html = html & "</td></tr>"

	end if

		html = html & "<tr>"
		html = html & "<td class='fil'>TOTAL (hs)</td><td class='fil'>" & rs(32) & "</td>"
		html = html & "</tr>"

	if rs(24) = session("USUARIO") and session("FUNC") = "FUNC" then

		html = html & "<tr>"
		html = html & "<td class='fil'>BROWSER</td><td class='fil'>"
		html = html & armaCombo("select id,descripcion from browsers",Cnn,rs(29),0,1,"browser")
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>VIA PEDIDO</td><td class='fil'>"
		html = html & armaCombo("select id,descripcion from via_pedido",Cnn,rs(30),0,1,"via_pedido")
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>REQUIERE PRESUPUESTO</td><td class='fil'>"
		html = html & "<select class='selec' style='width:90%;' id='req_presup' name='req_presup'>"

		If rs(20) = "SI" Then

			html = html & "<option value='SI' selected>SI</option>"
			html = html & "<option value='NO'>NO</option>"

		Else

			html = html & "<option value='SI'>SI</option>"
			html = html & "<option value='NO' selected>NO</option>"

		End If

		html = html & "</select></td></tr>"

	else

		html = html & "<tr>"
		html = html & "<td class='fil'>BROWSER</td><td class='fil'>"
		html = html & rs(18)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>VIA PEDIDO</td><td class='fil'>"
		html = html & rs(19)
		html = html & "</td></tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>REQUIERE PRESUPUESTO</td><td class='fil'>"
		html = html & rs(20) & "</td></tr>"

	end if

	if rs(25) = session("USUARIO") and session("FUNC") = "TEC" then

		html = html & "<tr>"
		html = html & "<td class='fil'>ESTADO DESARROLLO</td><td class='fil'>"
		html = html & armaCombo("select id,descripcion from estados",Cnn,rs(27),0,1,"estado_desa")
		html = html & "</td></tr>"

	else

		html = html & "<tr>"
		html = html & "<td class='fil'>ESTADO DESARROLLO</td><td class='fil'>"
		html = html & rs(11)
		html = html & "</td></tr>"

	end if

	if rs(24) = session("USUARIO") and session("FUNC") = "FUNC" then

		if rs(11) & "" = "OK" OR rs(11) & "" = "BAJA" then

			html = html & "<tr>"
			html = html & "<td class='fil'>ESTADO FUNCIONAL</td><td class='fil'>"
			html = html & armaCombo("select id,descripcion from estados",Cnn,rs(26),0,1,"estado_func")
			html = html & "</td></tr>"
			html = html & "<tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>ESTADO FUNCIONAL</td><td class='fil'>"
			html = html & rs(12)
			html = html & "</td></tr>"
			html = html & "<tr>"

		end if

		if rs(11) & "" = "OK" and rs(12) & "" = "OK" then

			html = html & "<tr>"
			html = html & "<td class='fil'>FECHA OK CLIENTE</td><td class='fil'>"
			html = html & "<input id='ok_cliente' name='ok_cliente' type='text' style='height:8%;width:100%;color:#000000;border:solid 1px #000000;' value='" & rs(33) & "'>"
			html = html & "</td></tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>FECHA OK CLIENTE</td><td class='fil'>"
			html = html & rs(33)
			html = html & "</td></tr>"

		end if

	else

		html = html & "<tr>"
		html = html & "<td class='fil'>ESTADO FUNCIONAL</td><td class='fil'>"
		html = html & rs(12)
		html = html & "</td></tr>"
		html = html & "<tr>"

		html = html & "<tr>"
		html = html & "<td class='fil'>FECHA OK CLIENTE</td><td class='fil'>"
		html = html & rs(33)
		html = html & "</td></tr>"

	end if

		html = html & "<tr>"
		html = html & "<td class='fil'>LINK DOCUMENTO</td><td class='fil'>"
		test = "<a href='" & rs("ruta") & "'><font style='color:#FF0000;'>Test</font></a>"
		html = html & "<input class='inpu' style='width:90%' id='ruta' name='ruta' size=50 type='text' value='" & rs("ruta") & "'>" & test
		html = html & "</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td colspan=2>&nbsp;</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td id='sAgre' class='fil' colspan=2 align=center><font style='color:#FF0000;'>RELACIONAR APL</font></td>"
		html = html & "</tr>"

		html = html & "</table>"

		html = html & "</div>"

		html = html & "<div style='height:100%;width:50%;float:right'><br>"

		html = html & "<table class='tablDet' style='width:80%;' cellspacing=0 align=left>"

		html = html & "<tr>"
		html = html & "<td class='fil' style='font-size:12pt;' ><b><u>DETALLES</u></b></td>"
		html = html & "</tr>"

	if rs(24) = session("USUARIO") and session("FUNC") = "FUNC" then

		html = html & "<tr>"
		html = html & "<td class='fil' >"
		html = html & "<textarea class='txtar' id='descripcion' name='descripcion' type='text' style='height:200px;width:100%;'>" & rs(3) & "</textarea>"
		html = html & "</td></tr>"

	else

		html = html & "<tr>"
		html = html & "<td class='fil'>"
		html = html & rs(3)
		html = html & "</td></tr>"

	end if

		html = html & "</table><br><br><br><br><br><br><br>"

		html = html & "<table class='tablDet' style='margin-top:50px;width:80%;' cellspacing=0 align=left>"

		html = html & "<tr>"
		html = html & "<td class='fil' style='font-size:12pt;' ><b><u>AVANCES</u></b></td>"
		html = html & "</tr>"

		if (rs(24) = session("USUARIO") and session("FUNC") = "FUNC") or (rs(25) = session("USUARIO") and session("FUNC") = "TEC") then

			html = html & "<tr>"
			html = html & "<td class='fil'>"
			html = html & "<textarea class='txtar' id='avances' name='avances' type='text' style='height:200px;width:100%;'>" & rs(35) & "</textarea>"
			html = html & "</td>"
			html = html & "</tr>"

		else

			html = html & "<tr>"
			html = html & "<td class='fil'>" & rs(35) & "</td>"
			html = html & "</tr>"

		end if

		html = html & "<tr>"
		html = html & "<td>&nbsp;</td>"
		html = html & "</tr>"

		html = html & "<tr>"
		html = html & "<td align=left>"
		html = html & "<input class='bot' type='button' value='ACTUALIZAR' onclick=""actualizar('sForm');"">"
		html = html & "</td></tr>"

		end if

		html = html & "</table>"

		html = html & "</form>"

		rs.close

	End If

	html = html & "</div>"

	html = html & "<script type='text/javascript'>"
	html = html & " function actualizar(formu){ "
	html = html & " var r = confirm('Desea actualizar ?'); "
	html = html & " if (r){ "
	html = html & " document.getElementById(formu).submit(); "
	html = html & " }else{ "
	html = html & " return false;"
	html = html & " }} "
	html = html & "</script>"

Else

	If opc = "veri" Then

		vuel = request.querystring("vuel")

		html = html & "<div>"

		html = html & "<header id='header' style='height:16%'>"

		html = html & "<div>"

		html = html & "<a href='index.asp' class='logo'>"

		html = html & "<img src='images/logocorP4.gif' width=10% >"
		html = html & "<font style='font-size:40pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"

		html = html & "</a>"

		html = html & "</div>"

		html = html & "</header>"

		html = html & "</div>"

		html = html & "<div><br><br>"

		html = html & "<table style='height:50%;width:90%;' align=center>"

		html = html & "<tr>"

		html = html & "<td style='height:100%;width:100%;' ><font style='color:#000000;font-size:30pt;'>INCIDENTE NO ENCONTRADO</font></td>"

		html = html & "</tr>"

		html = html & "<tr>"

		html = html & "<td style='height:100%;width:100%;' >"
		html = html & "<form id='xfrm' method='post' action='consultas.asp?sel=" & vuel & "' >"
		html = html & "<input style='border-radius:10px;border:solid 2px;border-color:blue;height:100%;width:50%;' value='VOLVER' type='button' onclick=""document.getElementById('xfrm').submit();"">"
		html = html & "</form>"
		html = html & "</td>"

		html = html & "</tr>"

		html = html & "</table>"

		html = html & "</div>"

	ElseIf opc = "verp" Then

		vuel = request.querystring("vuel")

		html = html & "<div>"

		html = html & "<header id='header' style='height:16%'>"

		html = html & "<div>"

		html = html & "<a href='index.asp' class='logo'>"

		html = html & "<img src='images/logocorP4.gif' width=10% >"
		html = html & "<font style='font-size:40pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"

		html = html & "</a>"

		html = html & "</div>"

		html = html & "</header>"

		html = html & "</div>"

		html = html & "<div><br><br>"

		html = html & "<table style='height:50%;width:90%;' align=center>"

		html = html & "<tr>"

		html = html & "<td style='height:100%;width:100%;' ><font style='color:#000000;font-size:30pt;'>PROYECTO NO ENCONTRADO</font></td>"

		html = html & "</tr>"

		html = html & "<tr>"

		html = html & "<td style='height:100%;width:100%;' >"
		html = html & "<form id='xfrm' method='post' action='consultas.asp?sel=" & vuel & "' >"
		html = html & "<input style='border-radius:10px;border:solid 2px;border-color:blue;height:100%;width:50%;' value='VOLVER' type='button' onclick=""document.getElementById('xfrm').submit();"">"
		html = html & "</form>"
		html = html & "</td>"

		html = html & "</tr>"

		html = html & "</table>"

		html = html & "</div>"

	End If


End If

html = html & "<div id='dialog' title='' >"
html = html & "</div>"

html = html & "<script>"
html = html & "  $( function() {"

html = html & "$.datepicker.regional['es'] = {"
html = html & "closeText: 'Cerrar',"
html = html & "prevText: '< Ant',"
html = html & "nextText: 'Sig >',"
html = html & "currentText: 'Hoy',"
html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
html = html & "dayNamesShort: ['Dom','Lun','Mar','Mié','Juv','Vie','Sáb'],"
html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
html = html & "weekHeader: 'Sm',"
html = html & "dateFormat: 'dd/mm/yy',"
html = html & "firstDay: 1,"
html = html & "isRTL: false,"
html = html & "showMonthAfterYear: false,"
html = html & "yearSuffix: ''"
html = html & "};"
html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"

html = html & "    $( ""#fecha_ingreso"" ).datepicker();"
html = html & "    $( ""#fecha_entrega"" ).datepicker();"
html = html & "    $( ""#ok_cliente"" ).datepicker();"
html = html & "  } );"
html = html & "  </script>"

html = html & "<script>"

html = html & "$( '#tabs' ).tabs();"

html = html & "$( '#dialog' ).dialog({"
html = html & "title: 'Agregar relacion',"
html = html & "autoOpen: false,"
html = html & "resizable: false,"
html = html & "buttons: {"
html = html & "'Aceptar': function () {"
html = html & "$.get('agreParam.asp?param=APL&dato1=' + $('#valor1').val() + '&dato2=' + $('#nid').text() + '&dato3=' + '" & opc & "', function( data ) {"
html = html & "});"
html = html & "alert('Agregado exitosamente');"
html = html & "$(this).dialog('close');"
html = html & "},"
html = html & "'Cerrar': function () {"
html = html & "$(this).dialog('close');"
html = html & "}"
html = html & "},"
html = html & "show: {"
html = html & "effect: 'blind',"
html = html & "duration: 100"
html = html & "      },"
html = html & "hide: {"
html = html & "effect: 'blind',"
html = html & "duration: 100"
html = html & "}"
html = html & "});"

html = html & "$( '#sAgre' ).click(function() {"
html = html & "$.get('armaCargaParam.asp?param=APL', function( data ) {"
html = html & "$( '#dialog' ).html(data);"
html = html & "});"
html = html & "$( '#dialog' ).dialog( 'open' );"
html = html & "});"

html = html & "</script>"

html = html & sformz & form7 & xform7 & form5 & xform7 & form8 & formz1 & "</body>"
html = html & "</html>"

Function armaCombo(sql,cn,clave,valor,texto,comboid)

	Set nrs = cn.Execute(sql)

	html = html & "<select class='selec' style='font-size:9pt;width:90%;color:#000000;border:solid 1px #000000;' id='" & comboid & "' name='" & comboid & "'>"

		if comboid <> "estado_desa" and comboid <> "estado_func" then
			html = html & "<option value='0'>S/D</option>"
		end if
	Do While Not nrs.eof

		If cstr(clave) = cstr(nrs(valor)) Then

			html = html & "<option value='" & nrs(valor) & "' selected>" & nrs(texto) & "</option>"

		Else

			html = html & "<option value='" & nrs(valor) & "'>" & nrs(texto) & "</option>"

		End If

		nrs.movenext
	Loop

	armarCombo = html & "</select>"

End Function

Function puedeAct(cn,id)

oSql = "select respon_gestion,respon_tec from incidentes where id=" & id
Set fcs = cn.Execute(oSql)

sPuede = "NO"
if fcs(0) = session("USUARIO") or fcs(1) = session("USUARIO") then sPuede = "SI"

puedeAct = sPuede

End Function

response.write html
%>