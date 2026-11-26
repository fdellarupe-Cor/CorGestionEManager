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
'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form
  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

opc = request.querystring("opc")

if opc = "cont" then

sForm = "sForm"
dest = request.querystring("dest")

'CODIGOS SELECCIONADOS
session("sIDs") = ""
delim = ""

For Each x In request.form
	if request.form(x) <> "adj" and request.form(x) <> "url" and request.form(x) <> "cor" then
	codigo = request.form(x)
	session("sIDs") = session("sIDs") & delim & codigo

	delim = ","
	end if
Next

if request.querystring("ids") & "" <> "" then session("sIDs") = request.querystring("ids")

if session("sIDs") = "" then session("sIDs") = 0

envio = request.form("checkEnv")
if envio & "" = "" then envio = request.querystring("checkEnv")
titEnvio = ""

select case envio

		case "adj"
			titEnvio = "Adjunto"
			dest = 1

		case "url"
			titEnvio = "URL"

		case "cor"
			titEnvio = "CorExpress"
			dest = 1

		case else
			titEnvio = "<font style='color:#FF0000;'>&nbsp;&nbsp;DEBE SELECCIONAR UN TIPO DE ENVIO</font>"

end select

'RUTINA PARA SACAR SOLO CODIGOS PARA MOSTRAR TABLA CON LOS DOC SELECCIONADOS
codigos = split(session("sIDs"),",")
cobjs = ""
delim = ""

for i=0 to ubound(codigos)

sdato = split(codigos(i),"-")

cobjs = cobjs & delim & sdato(0)
delim = ","

next

'TABLA FORMULARIOS SELECCIONADOS
sql = "select a.codobjeto,a.descri,e.archivo,c.fecpublic from objetos a "
sql = sql & " inner join tipobjeto b on (a.tipobjeto=b.tipobjeto and b.tipocpt='K') "
sql = sql & " inner join objprty c on (a.codobjeto=c.codobjeto and a.tipobjeto=c.tipobjeto and c.estado=3) "
sql = sql & " inner join objpropiedades d on (a.codobjeto=d.codobjeto and a.tipobjeto=d.tipobjeto) "
sql = sql & " inner join Objhtm e on (a.codobjeto=e.codobjeto and a.tipobjeto=e.tipobjeto) "
sql = sql & " where a.codobjeto in (" & cobjs & ")"

Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

if not rs.eof then

Celda = ofv.gencelda("","tt colspan=2","","","","","","Tipo Envio: " & titEnvio,"si")
Fila = ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","a colspan=2","","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt colspan=2","center","","","","","Formularios seleccionados","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt","","","","","","Descripcion","si")
Celda = Celda & ofv.gencelda("","tt","","","","","","Documento","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

cont = 0

do while not rs.eof

	Celda = ofv.gencelda("","a style='border:solid 1px;'","","","","","",rs(1),"si")
	Celda = Celda & ofv.gencelda("","a style='border:solid 1px;'","","","","","",rs(2),"si")
	Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

	cont = cont + 1

rs.movenext
loop

Celda = ofv.gencelda("","a colspan=2","","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt colspan=2","","","","","","Total documentos: " & cont,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


sHTML = sHTML & ofv.GenTabla("","a ","center","60%%","","0","","0","0","","",Fila,"si") & "<br>"

'TABLA DESTINOS solo para envio URL
if envio = "url" then

Celda = ofv.gencelda("","tt colspan=5","","","","","","Destinos","si")
Fila = ofv.genrow("","","","","","","",Celda,"si")

sql = "select id,destino from destinosurl where estado = 3 order by id"
Set rc = ofv.crearconsultaEx(sql,cn,1,parametros)

cont = 0
Celda = ""

do while not rc.eof

cont = cont + 1

if request.form("check" & rc(0)) & "" <> "" then
	chkbox = "<input id='check" & rc(0) & "' name='check" & rc(0) & "' class='dt' type='checkbox' value='" & rc(0) & "' checked>"
else
	chkbox = "<input id='check" & rc(0) & "' name='check" & rc(0) & "' class='dt' type='checkbox' value='" & rc(0) & "'>"
end if

Celda = Celda & ofv.gencelda("","","","12%","","","",chkbox & rc(1),"si")


if cont = 5 then

Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
cont = 0
Celda = ""

end if

rc.movenext
loop

if cont < 5 and cont > 0 then

Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
cont = 0

end if

Celda = ofv.gencelda("","a colspan=5","center","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'onclk = "onclick=" & chr(34) & "document.getElementById('" & sForm & "').submit();" & chr(34)
'btnSel = ofv.GenerarInput("","Crear lote","button","","10","bt " & onclk,"")
'Celda = ofv.gencelda("","a colspan=5","center","","","","",btnSel,"si")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


xTabla = ofv.GenTabla("","a ","center","60%%","","0","","0","0","","",Fila,"si") & "<br>"

end if

xsql = "select max(id) from envioformmaster"
Set rl = ofv.crearconsultaEx(xsql,cn,1,parametros)

maxId = ""
if not rl.eof then maxId = rl(0) + 1

sAccion = "formEnvDet.asp?opc=crear&dest=1&checkEnv=" & envio
sHTML = sHTML & ofv.GenForm(sForm,"",sAccion,"","post","",xTabla,"si")

if envio <> "url" then
	onclk = "onclick=" & chr(34) & "document.getElementById('" & sForm & "').submit();" & chr(34)
else
	onclk = "onclick=" & chr(34) & "validar('" & sForm & "',true);" & chr(34)
end if
sHTML = sHTML & "<center>" & ofv.GenerarInput("","Crear lote " & maxId,"button","","10","bt " & onclk,"") & "</center>"



'TABLA DESTINATARIOS
'if envio = "adj" or envio = "cor" or dest = 1 then

'Celda = ofv.gencelda("","tt","","","","","","DESTINATARIOS","si")
'Fila = ofv.genrow("","","","","","","",Celda,"si")

'if envio = "url" then

'	idx = ""
'	delimx = ""

'	For Each x In request.form
'		idx = idx & delimx & request.form(x)
'		delimx = ","
'	Next

'	if idx & "" = "" then idx = 0

'	sql = "select mail1,mail2 from destinosurl where id in (" & idx & ")"
'	Set rm = ofv.crearconsultaEx(sql,cn,1,parametros)

'	mails = ""
'	sep = ""
'	do while not rm.eof

'		mails = mails & sep & rm(0) & ";" & rm(1)
'		sep = ";"

'	rm.movenext
'	loop

'	sInput = "<textarea style='height:100%;width:100%;' id='dest' name='dest' cols=10 rows=5 >" & mails & "</textarea>"

'else

'	sInput = "<textarea style='height:100%;width:100%;' id='dest' name='dest' value='' cols=10 rows=5 ></textarea>"

'end if


'Celda = ofv.gencelda("","tt","","","","","",sInput,"si")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'btnEnv = ofv.GenerarInput("","Crear lote","submit","","10","bt","")
'Celda = ofv.gencelda("","","center","","","","",btnEnv,"si")
'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'pTabla = ofv.GenTabla("","a ","center","70%%","","0","","0","0","","",Fila,"si")

'sAccion = "formEnvDet.asp?opc=crear&envio=" & envio
'sHTML = sHTML & ofv.GenForm(sForm,"",sAccion,"","post","",pTabla,"si")




'end if


else

Celda = ofv.gencelda("","","center","","","","","<font style='color:#FF0000;font-size:15;'>NINGUN DOCUMENTO SELECCIONADO</font>","si")
Fila = ofv.genrow("","","","","","","",Celda,"si")

sHTML = sHTML & ofv.GenTabla("","a ","center","70%%","","0","","0","0","","",Fila,"si")

end if




ELSEIF opc = "det" then

cobj = request.querystring("cobj")
tobj = request.querystring("tobj")

sql = "select a.descri,b.archivo,c.fecenvio,d.tipoenv from objetos a "
sql = sql & " inner join objhtm b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto) "
sql = sql & " inner join envioformdet1 c on (a.codobjeto=c.codobjeto and a.tipobjeto=c.tipobjeto) "
sql = sql & " inner join envioformmaster d on (c.idlote=d.id) "
sql = sql & " where a.codobjeto=" & cobj & " and a.tipobjeto='" & tobj & "'"

Set rz = ofv.crearconsultaEx(sql,cn,1,parametros)

if not rz.eof then

Celda = ofv.gencelda("","tt colspan=2","center","","","","","Detalle envio","si")
Fila = ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","ttt","","20%","","","","Descripcion:","si")
Celda = Celda & ofv.gencelda("","ttt","","","","","",rz(0),"si")

Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","ttt","","20%","","","","Docum asociado:","si")
Celda = Celda & ofv.gencelda("","ttt","","","","","",rz(1),"si")

Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","ttt","","20%","","","","Fecha envio:","si")
Celda = Celda & ofv.gencelda("","ttt","","","","","",rz(2),"si")

Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","ttt","","20%","","","","Tipo envio:","si")

select case rz(3)
	case 1
		tEnv = "Adjunto"
	case 2
		tEnv = "Url"
	case 3
		tEnv = "Corexpress"
end select

Celda = Celda & ofv.gencelda("","ttt","","","","","",tEnv,"si")

Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

sHTML = sHTML & ofv.GenTabla("","","center","80%","","0","","0","0","","",Fila,"si")

end if


elseif opc = "crear" then

'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form & "<br>" & "IDs: " & session("sIDs")

nFecini = CLng(CDate("01-01-1900"))

sFecCrea = clng(date()) - nFecini

env = request.querystring("checkenv")

select case env

		case "adj"
			tEnv = 1

		case "url"
			tEnv = 2

		case "cor"
			tEnv = 3

end select

'INSERT EN LOTE MASTER
sql = "insert into envioformmaster (tipoenv,feccreac,fecenvio) values ('" & tEnv & "','" & sFecCrea & "','0')"
call ofv.crearconsultaEx(sql,cn,1,parametros)

'PIDO ULTIMO ID LOTE
ultId = getUltIdLote()

'INSERTO EN LA DETALLES1 TANTAS ENTRADAS COMO DOCUMENTOS HAYA SELECCIONADO
docs = split(session("sIDs"),",")

for i=0 to ubound(docs)

	xdatos = split(docs(i),"-")
	sql = "insert into envioformdet1 (idlote,codobjeto,tipobjeto,estado,fecenvio) values "
	sql = sql & "('" & ultId & "','" & xdatos(0) & "','" & xdatos(1) & "','NO','0')"
	call ofv.crearconsultaEx(sql,cn,1,parametros)

next

'INSERTO EN LA TABLA DETALLES2 SOLO SI ES TIPO URL PARA GUARDAR LOS ID DE LOS DESTINOS SELECCIONADOS
if tEnv = 2 then

For Each x In request.form

	sql = "insert into envioformdet2 (idlote,iddestino) values ('" & ultId & "','" & request.form(x) & "')"
	call ofv.crearconsultaEx(sql,cn,1,parametros)

Next

end if

response.redirect "EnvFormHist.asp"

END IF

Agregar = ""
Buscar = ""

sAccion = "formEnv.asp?" & request.querystring
Celda   = ofv.BotonVolver(sAccion)

sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

end if

sHTML = sHTML & "<script type='text/javascript'>"

sHTML = sHTML & "function validar(fid,dato)" & chr(13)
sHTML = sHTML & " {" & chr(13)
sHTML = sHTML & " var checkboxes = new Array(); var entro = 0; " & chr(13)
sHTML = sHTML & " checkboxes = document.getElementById(fid).getElementsByTagName('input'); " & chr(13)
sHTML = sHTML & " for (var i=0; i<checkboxes.length; i++) { " & chr(13)
sHTML = sHTML & " if (checkboxes[i].type == 'checkbox' & checkboxes[i].checked){entro=1;document.getElementById(fid).submit();} " & chr(13)
sHTML = sHTML & " } if (entro== 0){alert('Debe seleccionar al menos un destino.');}" & chr(13)
sHTML = sHTML & " }" & chr(13)

sHTML = sHTML & "</script>"


sHTML = sHTML & xFrmLimp & "</body></html>"

Response.Write sHTML

Function getUltIdLote()

sql = "select id from envioformmaster order by id desc"
set rl = ofv.crearconsultaEx(sql,cn,1,parametros)

xxID = ""
if not rl.eof then xxID = rl(0)

getUltIdLote = xxID

End Function

%>