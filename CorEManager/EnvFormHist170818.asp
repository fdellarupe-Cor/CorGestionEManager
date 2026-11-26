<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWkfManager.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrar = 10
  CantRegMover = CDbl("0" &Request.querystring("CantRegMover"))

  opc = request.querystring("opc")
  id = request.querystring("id")
  tEnv = request.querystring("tenv")

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_opc=" & opc & "_id=" & id & "_tenv=" & tEnv

  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form
  set cn = ofv.conectar(ofv.strconn2)

'LISTADO DE TODOS LOS LOTES
if opc & "" = "" then

ixForm = "ixForm"

bID = request.form("xID")
if bID & "" = "" then bID = request.querystring("xID")

bCF = request.form("xCF")
if bCF & "" = "" then bCF = request.querystring("xCF")

Celda = ofv.gencelda("","tt colspan=2","center","","","","","Filtro lotes","")
Fila = ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt","","","","","","Id","")
inputID = ofv.GenerarInput("xID",bID,"text",ixForm,"1","dt onkeypress='return isNumber(event)'","no")
Celda = Celda & ofv.gencelda("","tt","","","","","",inputID,"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt","","","","","","Cod form","")
inputCF = ofv.GenerarInput("xCF",bCF,"text",ixForm,"1","dt","no")
Celda = Celda & ofv.gencelda("","tt","","","","","",inputCF,"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

btnBus = ofv.GenerarInput("","Buscar","submit","","10","bt","")
Celda = ofv.gencelda("","a colspan=2","center","","","","",btnBus,"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

tabla = ofv.GenTabla("","a","center","20%","","0","","0","0","","",Fila,"si")

sAccion = "EnvFormHist.asp?opc="
sHTML = sHTML & ofv.GenForm(ixForm,"",sAccion,"","post","",tabla,"si")

'ENCABEZADO
Celda = ofv.gencelda("","tt","center","2%","","","","*","")
Celda = Celda & ofv.gencelda("","tt","","5%","","","","Id","")
Celda = Celda & ofv.gencelda("","tt","","10%","","","","Tipo envio","")
Celda = Celda & ofv.gencelda("","tt","","10%","","","","Fecha creacion","")
Celda = Celda & ofv.gencelda("","tt","center","3%","","","","*","")
Celda = Celda & ofv.gencelda("","tt","center","3%","","","","*","")

Formulario = ofv.genrow("","","","","","","",Celda,"si")

sql = "select distinct a.* from envioformmaster a "
sql = sql & " left join envioformdet1 b on (a.id=b.idlote) "
sql = sql & " left join objpropiedades c on (b.codobjeto=c.codobjeto and b.tipobjeto=c.tipobjeto) "
sql = sql & " where a.id > 0 "

if bID & "" <> "" then
	sql = sql & " and a.id=" & bID
end if

if bCF & "" <> "" then
	sql = sql & " and c.codigoform='" & bCF & "'"
end if

sql = sql & " order by id desc"
'response.write sql
Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

style = "style='border-bottom:solid 1px;'"

xTrColor = "#FFFFFF"

PosReg = 0

if not rs.eof then
rs.Move (CantRegMover)
Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

'fecha creacion
nFecini = CLng(CDate("01-01-1900"))
sFec = FechadorCor(cdate(rs(2) + nFecini),"")

'boton enviar
sAccion = "EnvFormHist.asp?opc=enviar&id=" & rs(0) & "&tenv=" & rs(1) & "&xID=" & bID & "&xCF=" & bCF
sHTML = sHTML & ofv.GenForm("sFormE" & rs(0),"",sAccion,"","post","",xTabla,"si")
onclick = "onclick=" & chr(34) & "document.getElementById('sFormE" & rs(0) & "').submit();" & chr(34)
btnEnv = ofv.GenerarInput("","Enviar","button","","10","bt " & onclick,"")

'boton detalle
sAccion = "EnvFormHist.asp?opc=deta&id=" & rs(0) & "&tenv=" & rs(1) & "&xID=" & bID & "&xCF=" & bCF
sHTML = sHTML & ofv.GenForm("sFormD" & rs(0),"",sAccion,"","post","",xTabla,"si")
onclick = "onclick=" & chr(34) & "document.getElementById('sFormD" & rs(0) & "').submit();" & chr(34)
btnDet = ofv.GenerarInput("","Detalle","button","","10","bt " & onclick,"")

'boton eliminar
if bSuprimir then
	sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
	sHref = sHref & "&Tabla=envioformmaster&volver=" & vuelta & "&Eliminar=NSNC"
	Celda = ofv.BotonEliminar(rs(0),sHref)
else
	Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
end if

tEnv = rs(1)

scomK = 1 & chr(9) & 2 & chr(9) & 3
scomD = "Adjunto" & chr(9) & "Url" & chr(9) & "CorExpress"

comboTipEnv = ofv.gencomboFX("lForm" & rs(0),"tipoenv",tEnv,scomK,scomD,"")

Celda = Celda & ofv.gencelda("","a " & style,"center","","","","",rs(0),"")
Celda = Celda & ofv.gencelda("","a " & style,"","","","","",comboTipEnv,"")
Celda = Celda & ofv.gencelda("","a " & style,"center","","","","",sFec,"")
if rs(3) <> 0 then
	Celda = Celda & ofv.gencelda("","","","","","","",btnDet,"")
else
	Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","")
end if
Celda = Celda & ofv.gencelda("","","","","","","",btnEnv,"")

Fila = ofv.genrow("","aaa bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

if xTrColor = "#FFFFFF" then
	xTrColor = "#F0F0F0"
ELSE
	xTrColor = "#FFFFFF"
END IF

sAccion = ""
if bModificar then
   sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
   sAccion = sAccion & "&tabla=envioformmaster&volver=" & vuelta
end if

Formulario = Formulario & ofv.genform("lForm" & rs(0),"",sAccion,"","post","",Fila,"si")

PosReg = PosReg + 1

rs.movenext
loop

sHTML = sHTML & ofv.GenTabla("","","center","33%","","0","","0","0","","",Formulario,"si")

end if

ELSEIF opc = "deta" then

envio = ""
select case tEnv

		case 1
			envio = "Adjunto"

		case 2
			envio = "Url"

		case 3
			envio = "CorExpress"

end select

sql = "select a.*,b.codobjeto,b.tipobjeto,b.fecenvio,d.descri,b.objid from envioformmaster a "
sql = sql & " left join envioformdet1 b on (a.id=b.idlote) "
sql = sql & " left join objetos d on (b.codobjeto=d.codobjeto and b.tipobjeto=d.tipobjeto) "
sql = sql & " where a.id=" & id

Set rg = ofv.crearconsultaEx(sql,cn,1,parametros)

nFecini = CLng(CDate("01-01-1900"))
sFecCreac = FechadorCor(cdate(rg(2) + nFecini),"")

sFecEnvio = ""
if rg(3) <> 0 then
	sFecEnvio = FechadorCor(cdate(rg(3) + nFecini),"")
else
	sFecEnvio = "No enviado"
end if

Celda = ofv.gencelda("","tt","","15%","","","","ID: " & rg(0),"")
Celda = Celda & ofv.gencelda("","tt","","15%","","","","Tipo Envio: " & envio,"")
Celda = Celda & ofv.gencelda("","tt","","15%","","","","Fecha creacion: " & sFecCreac,"")
Celda = Celda & ofv.gencelda("","tt","","15%","","","","Fecha envio: " & sFecEnvio,"")

Fila = ofv.genrow("","","","","","","",Celda,"si")

sHTML = sHTML & ofv.GenTabla("","","center","60%","","0","","0","0","","",Fila,"si")

Celda = ofv.gencelda("","a colspan=5","","","","","","&nbsp;","")
Fila = ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt colspan=5","center","","","","","Detalle envios","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt","","1%","","","","*","")
Celda = Celda & ofv.gencelda("","tt colspan=2","","","","","","Descripcion","")
Celda = Celda & ofv.gencelda("","tt colspan=2","","","","","","Fecha y hora envio","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

if rg(8) & "" <> "" then

cont = 0

do while not rg.eof

'boton eliminar
if bSuprimir and esUltimoDoc(rg(0)) = "NO" then
	sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rg(8)," ","_")
	sHref = sHref & "&Tabla=envioformdet1&volver=" & vuelta & "&Eliminar=NSNC"
	Celda = ofv.BotonEliminar(rg(8),sHref)
else
	Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
end if

xfec = ""
if rg(6) = "0" then
	xfec = "S/D"
else
	xfec = rg(6)
end if

Celda = Celda & ofv.gencelda("","a style='border:solid 1px;' colspan=2","","","","","",rg(7),"")
Celda = Celda & ofv.gencelda("","a style='border:solid 1px;' colspan=2","","","","","",xfec,"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

cont = cont + 1

rg.movenext
loop


Celda = ofv.gencelda("","tt colspan=5","","","","","","Total documentos: " & cont,"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda(""," colspan=5","","","","","","&nbsp;","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

else

Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=5","","","","","","<font style='color:#FF0000;font-size:12pt;'>SIN DOCUMENTOS</font>","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda(""," colspan=5","","","","","","&nbsp;","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
end if

'if tEnv = 2 then

'sql = "select b.iddestino from envioformmaster a "
'sql = sql & " left join envioformdet2 b on (a.id=b.idlote) "
'sql = sql & " inner join destinosurl c on (b.iddestino=c.id) "
'sql = sql & " where a.id=" & id

'Set rp = ofv.crearconsultaEx(sql,cn,1,parametros)

'destinos = ""
'urls = ""
'mails = ""
'sep = ""

'do while not rp.eof

'	sql = "select destino,url,mail1,mail2 from destinosurl where id=" & rp(0)
'	Set rh = ofv.crearconsultaEx(sql,cn,1,parametros)

'	if not rh.eof then

'		destinos = destinos & sep & rh(0)
'		urls = urls & rh(1) & "<br>"
'		mails = mails & rh(2) & "<br>" & rh(3) & "<br>"

'		sep = ",&nbsp;"

'	end if

'rp.movenext
'loop


'Celda = ofv.gencelda("","tt colspan=4","center","","","","","Datos preseleccionados","")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'Celda = ofv.gencelda("","tt colspan=4","","","","","","Destinos","")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'if destinos & "" <> "" then
'	Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","",destinos,"")
'else
'	Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","","S/D","")
'end if
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
'
'Celda = ofv.gencelda("","tt colspan=4","","","","","","Urls","")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
'
'if urls & "" <> "" then
'	Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","",urls,"")
'else
	'Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","","S/D","")
'end if
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'Celda = ofv.gencelda("","tt colspan=4","","","","","","Mails","")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'if mails & "" <> "" then
'	Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","",mails,"")
'else
'	Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","","S/D","")
'end if
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'else

sql = "select mails from envioformdet3 where idlote=" & id
Set ri = ofv.crearconsultaEx(sql,cn,1,parametros)

Celda = ofv.gencelda("","tt colspan=4","","","","","","Destinatarios último envio","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

if not ri.eof then

Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","",replace(ri(0) & "" ,";","<br>"),"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

else

Celda = ofv.gencelda("","a style='border:solid 1px;' colspan=4","","","","","","S/D","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

end if

'end if

sHTML = sHTML & ofv.GenTabla("","","center","60%","","0","","0","0","","",Fila,"si")



'PANTALLA PREVIA AL PROCESO DE ENVIO
ELSEIF opc = "enviar" then

id = request.querystring("id")
tEnvio = request.querystring("tenv")

strsql = "select a.ID,a.tipoenv,a.feccreac,c.descri,d.archivo "
strsql = strsqL & " from envioformmaster a "
strsql = strsqL & " inner join envioformdet1 b on (a.id=b.idlote) "
strsql = strsqL & " inner join OBJETOS c on (b.codobjeto=c.codobjeto and b.tipobjeto=c.tipobjeto) "
strsql = strsqL & " inner join objhtm d on (b.codobjeto=d.codobjeto and b.tipobjeto=d.tipobjeto) "
strsql = strsqL & " where a.id=" & id & " and tipoenv=" & tEnvio

Set rx = ofv.crearconsultaEx(strsql,cn,1,parametros)

'CABECERA
tEnv = ""
select case rx(1)

		case 1
			tEnv = "Adjunto"

		case 2
			tEnv = "Url"

		case 3
			tEnv = "CorExpress"

end select

fi = CLng(CDate("01-01-1900"))
fCreac = CDate(rx(2)) + fi
fCreac = FechadorCor(fCreac,"")

Celda = ofv.gencelda("","tt","","10%","","","","ID: " & rx(0),"")
Celda = Celda & ofv.gencelda("","tt","","25%","","","","Envio: " & tEnv,"")
Celda = Celda & ofv.gencelda("","tt","","25%","","","","Fecha creacion: " & fCreac,"")

Fila = ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","a colspan=3","","10%","","","","&nbsp;","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

sHTML = sHTML & ofv.GenTabla("","","center","60%","","0","","0","0","","",Fila,"si")


'DOCUMENTOS
Celda = ofv.gencelda("","tt","","30%","","","","Documento","")
Celda = Celda & ofv.gencelda("","tt","","29%","","","","Archivo","")
Fila = ofv.genrow("","","","","","","",Celda,"si")

cont = 0

do while not rx.eof

Celda = ofv.gencelda("","a style='border: solid 1px;'","","28%","","","",rx(3),"")
Celda = Celda & ofv.gencelda(""," a style='border: solid 1px;'","","27%","","","",rx(4),"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

cont = cont + 1

rx.movenext
loop

Celda = ofv.gencelda("","a colspan=3","","","","","","&nbsp;","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt colspan=3","","","","","","Total documentos: " & cont,"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","a colspan=3","","","","","","&nbsp;","")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

sHTML = sHTML & ofv.GenTabla("","","center","60%","","0","","0","0","","",Fila,"si")

'DESTINATARIOS
'if tEnvio <> 2 then

	'Celda = ofv.gencelda("","tt","center","5%","","","","Destinatarios","")
	'Filax = ofv.genrow("","","","","","","",Celda,"si")
	'sInput = "<textarea style='height:100%;width:100%;' id='dest' name='dest' cols=10 rows=5 ></textarea>"
	'Celda = ofv.gencelda("","tt","","5%","","","",sInput,"")
	'Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

'else

	oSql = "select mails from envioformdet3 where idlote=" & id
	Set ma = ofv.crearconsultaEx(osql,cn,1,parametros)

	if ma.eof or tEnvio = 2 then

	sql = "select iddestino from envioformdet2 where idlote=" & id
	Set rm = ofv.crearconsultaEx(sql,cn,1,parametros)

	if not rm.eof then

	Celda = ofv.gencelda("","tt","center","5%","","","","Destinatarios","")
	Filax = ofv.genrow("","","","","","","",Celda,"si")

	dests = ""
	delim = ""
	do while not rm.eof

	dests = dests & delim & rm(0)
	delim = ","

	rm.movenext
	loop

	sql = "select destino,url,mail1,mail2 from destinosurl where id in(" & dests & ")"
	Set rv = ofv.crearconsultaEx(sql,cn,1,parametros)

	nomDest = ""
	urls = ""
	mails1 = ""
	mails2 = ""
	delim = ""
	delim2 = ""
	do while not rv.eof

		nomDest = nomDest & delim & rv(0)
		urls = urls & rv(1) & "<br>"

		if rv(2) & "" <> "" then
			mails1 = mails1 & delim2 & rv(2)
		end if

		if rv(3) & "" <> "" then
			mails2 = mails2 & delim2 & rv(3)
		end if

		delim = ","
		delim2 = ";"

	rv.movenext
	loop

	if tEnvio = 2 then

	Celda = ofv.gencelda("","tt","","5%","","","","Destinos preseleccionados","")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")
	Celda = ofv.gencelda("","a style='border:solid 1px;'","","5%","","","",nomDest,"")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

	Celda = ofv.gencelda("","tt","","5%","","","","URLs","")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")
	Celda = ofv.gencelda("","a style='border:solid 1px;'","","5%","","","",urls,"")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

	end if


	xMails = ""

	if mails1 & "" <> "" then
		xMails = mails1
	end if

	if mails2 & "" <> "" then
			xMails = xMails & ";" & mails2
	end if

	Celda = ofv.gencelda("","tt","","5%","","","","Mails","")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")
	sInput = "<textarea style='height:100%;width:100%;' id='dest' name='dest' cols=10 rows=5 ></textarea>"
	Celda = ofv.gencelda("","tt","","5%","","","",sInput,"")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

	else

	if tEnvio = 2 then

	Celda = ofv.gencelda("","tt colspan=5","","","","","","Destinos","si")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

	sql = "select id,destino from destinosurl where estado = 3 order by id"
	Set rc = ofv.crearconsultaEx(sql,cn,1,parametros)

	cont = 0
	Celda = ""

	do while not rc.eof

	cont = cont + 1

	'if request.form("check" & rc(0)) & "" <> "" then
	'	chkbox = "<input id='check" & rc(0) & "' name='check" & rc(0) & "' class='dt' type='checkbox' value='" & rc(0) & "' checked>"
	'else
		chkbox = "<input id='check" & rc(0) & "' name='check" & rc(0) & "' class='dt' type='checkbox' value='" & rc(0) & "'>"
	'end if

	Celda = Celda & ofv.gencelda("","","","12%","","","",chkbox & rc(1),"si")


	if cont = 5 then

	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")
	cont = 0
	Celda = ""

	end if

	rc.movenext
	loop

	if cont < 5 and cont > 0 then

	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")
	cont = 0

	end if

	Celda = ofv.gencelda("","a colspan=5","center","","","","","&nbsp;","si")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")


	end if

	if tEnvio = 2 then
		Celda = ofv.gencelda("","tt colspan=2","center","5%","","","","DESTINATARIOS","")
	else
		Celda = ofv.gencelda("","tt","center","5%","","","","DESTINATARIOS","")
	end if

	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")
	sInput = "<textarea style='height:100%;width:100%;' id='dest' name='dest' cols=10 rows=5 ></textarea>"

	if tEnvio = 2 then
		Celda = ofv.gencelda("","tt colspan=2","","5%","","","",sInput,"")
	else
		Celda = ofv.gencelda("","tt","","5%","","","",sInput,"")
	end if

	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

	end if

else


	Celda = ofv.gencelda("","tt","","5%","","","","Mails","")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")
	sInput = "<textarea style='height:100%;width:100%;' id='dest' name='dest' cols=10 rows=5 ></textarea>"
	Celda = ofv.gencelda("","tt","","5%","","","",sInput,"")
	Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

end if

'end if


Celda = ofv.gencelda("","","","5%","","","","&nbsp;","")
Filax = Filax & ofv.genrow("","","","","","","",Celda,"si")

onclk = "onclick=" & chr(34) & "document.getElementById('xForm').submit();" & chr(34)
btnEnviar = ofv.GenerarInput("","Enviar","button","","10","bt " & onclk,"")

if tEnvio = 2 then
	Celda = ofv.gencelda("","a colspan=2","","5%","","","",btnEnviar,"")
else
	Celda = ofv.gencelda("","","","5%","","","",btnEnviar,"")
end if

Filax = Filax & ofv.genrow("","","center","","","","",Celda,"si")

sTabla = ofv.GenTabla("","","center","60%","","0","","0","0","","",Filax,"si")

sAccion = "EnvFormHist.asp?opc=fin&id=" & id & "&tenv=" & tEnvio
sHTML = sHTML & ofv.GenForm("xForm","",sAccion,"","post","",sTabla,"si")


'PROCESO DE ENVIO
ELSEIF opc = "fin" then
'response.write request.querystring & "<br>" & request.form
id = request.querystring("id")
tEnvio = request.querystring("tenv")

sql = "select texto from textoenvio where tipoenv=" & tEnvio
Set ro = ofv.crearconsultaEx(sql,cn,1,parametros)

if not ro.eof then cuerpo = ro(0)

cuerpo = cuerpo & "<br><br>Lote: " & id

if session("APLMDCLI") = "BGAL" then
	asunto = "Claridad Galicia - Envio de formularios"
else
	asunto = "Envio de formularios"
end if

mailTo = request.form("dest")

sql = "select b.directorio,b.archivo,a.codobjeto,a.tipobjeto,c.descri,b.tipo from envioformdet1 a "
sql = sql & " inner join objhtm b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto) "
sql = sql & " inner join objetos c on (a.codobjeto=c.codobjeto and a.tipobjeto=c.tipobjeto) "
sql = sql & " where a.idlote=" & id

Set ri = ofv.crearconsultaEx(sql,cn,1,parametros)

select case tEnvio

		'ADJUNTO
		case 1

		do while not ri.eof

			cuerpo = ""
			doc = ""
			if not ro.eof then cuerpo = ro(0)

			cuerpo = cuerpo & "<br><br>Lote: " & id

			doc = "<br>Documento EManager:&nbsp;<a href='" & getobjurl(ri(2),ri(3)) & "'>" & ri(4) & "</a>"

			cuerpo = cuerpo & doc

			if ri(5) = 1 then
				adjunto = Request.ServerVariables("APPL_PHYSICAL_PATH") & "COREMANAGER\" & ri(0) & "\" & ri(1)
			elseif ri(5) = 3 then
				adjunto = ri(1)
			end if

			if session("APLMDCLI") = "BGAL" then
				sCopia = "xClaridadGalicia@bancogalicia.com.ar"
			else
				sCopia = ""
			end if

			xx = SendMailAtt(cuerpo,mailTo,"mabraham@corconsulting.com.ar",asunto,"CDOSYS","SI","1","NO",adjunto,sCopia,tEnvio)

			sql = "update envioformdet1 set fecenvio='" & NOW() & "',estado='SI' where idlote=" & id & " and codobjeto=" & ri(2) & " and tipobjeto='" & ri(3) & "'"
			call ofv.crearconsultaEx(sql,cn,1,parametros)

		ri.movenext
		loop

		'URL
		case 2

			if session("COPIA") = "S" then


				for each x in request.form
					if isNumeric(request.form(x)) then

						iSql = "insert into envioformdet2 (idlote,iddestino) values ('" & id & "','" & request.form(x) & "')"
						call ofv.crearconsultaEx(iSql,cn,1,parametros)

					end if
				next


				dim fs
				set fs=Server.CreateObject("Scripting.FileSystemObject")

				xMail = ""
				sep = ""
				xFlg = 0
				do while not ri.eof

					strsql = "select url,mail1,mail2 from destinosurl a "
					strsql = strsql & " inner join envioformdet2 b on (a.id = b.iddestino) "
					strsql = strsql & " where b.idlote=" & id & " and b.iddestino <> 0"
					set ry = ofv.crearconsultaEx(strsql,cn,1,parametros)

					do while not ry.eof

						if xFlg = 0 then
							xMail = xMail & sep & ry(1) & ";" & ry(2)
						end if

						if ri(5) = 1 then
							xOrigen = session("DirUpload") & ri(0) & "\" & ri(1)
						elseif ri(5) = 3 then
							xOrigen = ri(1)
						end if

						xDestino = ry(0)

						fs.CopyFile xOrigen,xDestino

						sep = ";"

					ry.movenext
					loop

					xFlg = 1

					sql = "update envioformdet1 set fecenvio='" & NOW() & "',estado='SI' where idlote=" & id & " and codobjeto=" & ri(2) & " and tipobjeto='" & ri(3) & "'"
					call ofv.crearconsultaEx(sql,cn,1,parametros)

					cuerpo = cuerpo & "<br>Documento EManager:&nbsp;<a href='" & getobjurl(ri(2),ri(3)) & "'>" & ri(4) & "</a>"

				ri.movenext
				loop

				if session("APLMDCLI") = "BGAL" then
					sCopia = "xClaridadGalicia@bancogalicia.com.ar"
				else
					sCopia = ""
				end if

				xx = SendMailAtt(cuerpo,mailTo,"mabraham@corconsulting.com.ar",asunto,"CDOSYS","SI","1","NO","",sCopia,tEnvio)

				set fs=nothing

		end if

		'COREXPRESS
		case 3

		URLS = ""
		cont = 1

		do while not ri.eof

			URLS = URLS & "<a href='" & getobjurl(ri(2),ri(3)) & "'>" & ri(4) & "</a><br>"

			sqlx = "update envioformdet1 set fecenvio='" & NOW() & "',estado='SI' where idlote=" & id & " and codobjeto=" & ri(2) & " and tipobjeto='" & ri(3) & "'"
			call ofv.crearconsultaEx(sqlx,cn,1,parametros)

			cont = cont + 1

		ri.movenext
		loop

		if session("APLMDCLI") = "BGAL" then
			sCopia = "xClaridadGalicia@bancogalicia.com.ar"
		else
			sCopia = ""
		end if

		xx = SendMailAtt(cuerpo & "<br><br>" & URLS,mailTo,"mabraham@corconsulting.com.ar",asunto,"CDOSYS","SI","1","NO","",sCopia,tEnvio)

		case else

end select

nFecini = CLng(CDate("01-01-1900"))
sFecEnvio = clng(date()) - nFecini

'response.write request.querystring & "<br>" & request.form

if tEnvio <> 0 then

	sql = "update envioformmaster set fecenvio='" & sFecEnvio & "' where id=" & id
	call ofv.crearconsultaEx(sql,cn,1,parametros)

	iSql = "select objid from envioformdet3 where idlote=" & id
	set xm = ofv.crearconsultaEx(isql,cn,1,parametros)
'response.write "<br>xmail" & xMail & "<br>mailto" & mailTo
	if xm.eof then

		if tEnvio = 2 then
			'sMail = xMail & ";" & mailTo
			sMail = replace(xMail,";;",";")

			if mailto & "" <> "" then
				sMail = sMail & ";" & mailTo
			end if

			xSsQl = "select objid from envioformdet2 where idlote=" & id
			set xp = ofv.crearconsultaEx(xSsQl,cn,1,parametros)

			if xp.eof then
				sql = "insert into envioformdet3 (idlote,tipoenv,mails) values ('" & id & "','" & tEnvio & "','" & sMail & "')"
			else
				sql = "insert into envioformdet3 (idlote,tipoenv,mails) values ('" & id & "','" & tEnvio & "','" & mailto & "')"
			end if

		else

			sql = "insert into envioformdet3 (idlote,tipoenv,mails) values ('" & id & "','" & tEnvio & "','" & mailTo & "')"

		end if
		call ofv.crearconsultaEx(sql,cn,1,parametros)

	else

		if tEnvio = 2 then
			'sMail = xMail & ";" & mailTo
			sMail = replace(xMail,";;",";")

			if mailto & "" <> "" then
				sMail = sMail & ";" & mailTo
			end if

			xSsQl = "select objid from envioformdet2 where idlote=" & id
			set xp = ofv.crearconsultaEx(xSsQl,cn,1,parametros)

			if xp.eof then
				sSql = "update envioformdet3 set mails='" & sMail & "' where idlote=" & id
			else
				sSql = "update envioformdet3 set mails='" & mailTo & "' where idlote=" & id
			end if

		else
			sSql = "update envioformdet3 set mails='" & mailTo & "' where idlote=" & id
		end if
		call ofv.crearconsultaEx(ssql,cn,1,parametros)

	end if

end if

response.redirect "EnvFormHist.asp?opc="

end if








Agregar = ""
Buscar = ""

xID = request.querystring("xID")
xCF = request.querystring("xCF")

if opc <> "" then
	sAccion = "EnvFormHist.asp?opc=&xID=" & xID & "&xCF=" & xCF
	Celda = ofv.BotonVolver(sAccion)
else
	Celda = ""
end if

if opc <> "" then
	sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Celda)
else
	sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)
end if

sHTML = sHTML &  "<script>function isNumber(evt) {"
sHTML = sHTML & "    evt = (evt) ? evt : window.event;"
sHTML = sHTML & "    var charCode = (evt.which) ? evt.which : evt.keyCode;"
sHTML = sHTML & "    if (charCode > 31 && (charCode < 48 || charCode > 57)) {"
sHTML = sHTML &     "    return false;"
sHTML = sHTML &    " }"
sHTML = sHTML &    " return true;"
sHTML = sHTML & "}</script>"

sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

Function esUltimoDoc(idLote)

sql = "select count(*) from envioformdet1 where idlote=" & idLote
set rj = ofv.crearconsultaEx(sql,cn,1,parametros)

esUlt = ""

if not rj.eof then

	if rj(0) =< 1 then
		esUlt = "SI"
	else
		esUlt = "NO"
	end if

end if

esUltimoDoc = esUlt

End function

%>