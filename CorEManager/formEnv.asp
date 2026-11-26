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
 ' response.write "FORM: " & request.form
  acc=request.querystring("acc")
  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

'XFNAME   = request.querystring("XFNAME")
'XFSRH    = request.querystring("XFSRH")
'XFSRD    = request.querystring("XFSRD")

'IF XFNAME = "" THEN XFNAME = "PV"
'IF XFSRD  = "" THEN XFSRD = 0
'IF XFSRH  = "" THEN XFSRH = 0

'PRIMERA VEZ
'IF XFNAME="PV"  THEN
'	XFSRD = clng(fechaconvert(DATE())) - clng(cdate("01-01-1900"))
'	XFSRH = XFSRD '+ 15
'END IF

'IF XFNAME = "FD" THEN
'	XFSRD = request.form("DESDE")
'	XFSRH = XFSRH
'END IF

'IF XFNAME = "FH" THEN
'	XFSRD = XFSRD
'	XFSRH =request.form("HASTA")
'END IF

sImagenE = "<B><FONT FACE=CorFlatIcon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"

sForm = "sForm"
xForm = "xForm"

'INPUT CODIGO FORM
cfrm = ""
if request.form("cform") & "" <> "" then cfrm = request.form("cform")
if cfrm & "" = "" then cfrm = request.querystring("cform")
inputCForm = ofv.GenerarInput("cform",cfrm,"text",sForm,"1","dt","no")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'COMBO PRODUCTO
sql = "select objid,descri from tipoproducto order by descri"
set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

scomK = "0"
scomD = "TODOS"
sep = chr(9)

producto = "0"

if request.form("producto") & "" <> "" then producto = request.form("producto")
if producto & "" = "" then producto = request.querystring("producto")

do while not rs.eof

	scomK = scomK & sep & rs(0)
	scomD = scomD & sep & rs(1)

rs.movenext
loop

comboProd = ofv.gencomboFX(sForm,"producto",producto,scomK,scomD,"no")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'COMBO BANCA
sql = "select objid,descripcion from FORMADMbanca order by descripcion"
set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

scomK = "0"
scomD = "TODOS"
sep = chr(9)

banca = "0"

if request.form("banca") <> "" then banca = request.form("banca")
if banca & "" = "" then banca = request.querystring("banca")

do while not rs.eof

	scomK = scomK & sep & rs(0)
	scomD = scomD & sep & rs(1)

rs.movenext
loop

combobanca = ofv.gencomboFX(sForm,"banca",banca,scomK,scomD,"no")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'COMBO PROCESO
sql = "select objid,descripcion from FORMADMPROCESO order by descripcion"
set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

scomK = "0"
scomD = "TODOS"
sep = chr(9)

proceso = "0"

if request.form("proceso") <> "" then proceso = request.form("proceso")
if proceso & "" = "" then proceso = request.querystring("proceso")

do while not rs.eof

	scomK = scomK & sep & rs(0)
	scomD = scomD & sep & rs(1)

rs.movenext
loop

comboProceso = ofv.gencomboFX(sForm,"proceso",proceso,scomK,scomD,"no")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'INPUT DESCRIPCION
desc = ""
if request.form("descri") & "" <> "" then desc = request.form("descri")
if desc & "" = "" then desc = request.querystring("descri")
inputDescri = ofv.GenerarInput("descri",desc,"text",sForm,"194","dt","no")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'INPUT FEC PUBLIC DESDE
fd = ""
if request.form("fdesde") & "" <> "" then fd = request.form("fdesde")
if fd & "" = "" then fd = request.querystring("fdesde")
inputFdesde = ofv.GenerarInput("fdesde",fd,"text",sForm,"1","dt","no")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'INPUT FEC PUBLIC HASTA
fh = ""
if request.form("fhasta") & "" <> "" then fh = request.form("fhasta")
if fh & "" = "" then fh = request.querystring("fhasta")
inputFhasta = ofv.GenerarInput("fhasta",fh,"text",sForm,"1","dt","no")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'BOTON CONSULTAR
btnCons = "<input type=button value=Buscar onclick=""ejecutar('" & sform &"');"" class=bt>" 'ofv.GenerarInput("","Buscar","submit","","10","bt","")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'BOTON LIMPIAR
onclk = "onclick=" & chr(34) & "document.getElementById('limpiar').submit();" & chr(34)
btnLimp = ofv.GenerarInput("","Limpiar","button","","10","bt " & onclk,"")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'FILA TITULO
Celda = ofv.gencelda("","tt colspan=4","center","","","","","<h3>FILTRO PARA FORMULARIOS</h3>","si")
Fila = ofv.genrow("","","","","","","",Celda,"si")

'FILA 1
Celda = ofv.gencelda("","ttt","","20%","","","","Codigo formulario","si")
Celda = Celda & ofv.gencelda("","ttt","","20%","","","",inputCForm,"si")
Celda = Celda & ofv.gencelda("","ttt","","20%","","","","Banca","si")
Celda = Celda & ofv.gencelda("","ttt","","20%","","","",comboBanca,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'FILA 2
Celda = ofv.gencelda("","ttt","","","","","","Producto","si")
Celda = Celda & ofv.gencelda("","ttt","","","","","",comboProd,"si")
Celda = Celda & ofv.gencelda("","ttt","","","","","","Proceso","si")
Celda = Celda & ofv.gencelda("","ttt","","","","","",comboProceso,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'FILA 3
Celda = ofv.gencelda("","ttt","","","","","","Descripcion","si")
Celda = Celda & ofv.gencelda("","ttt colspan=3","","","","","",inputDescri,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'FILA 4
Celda = ofv.gencelda("","ttt","","","","","","Fecha Publicacion Desde","")

'sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,15,2,0,0,0)
'sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
'sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

' CELDAS DE DESDE
'if  clng(XFSRD) = 0 then
'	sInput2 = "------------"
'else
'	sInput2 = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
'	swd = datepart("w",sInput2)
'	sInput2 = Fechador(XFSRD)
'END IF

'sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","no") & ofv.GenerarInput("Desde3",cstr(cdate(XFSRD + clng(cdate("01-01-1900")))),"hidden","","20","dt","no")
  inputFdesde = ofv.GenerarInput("fdesde",fd,"hidden",sForm,"50","dt","no")

		if isnumeric(fd) then


		      dFecha = cdate(fd + clng(cdate("01-01-1900")))
		      xfecact = fd
		   '   dFechan = Fechador(xfecact)
		      dFechan = GTFechador(fd)

	      else
	          dFecha = ""
			 xfecact = "N"
		      dFechan = "-------"

         end if

		      sImagenI4 = "<B><i class='las la-calendar' style=font-size:14pt;color:" & Session("CMFPNBOT") & "></i></B>"
		      sinput = "<a class=cc style='text-decoration:none;background-color:none;' xonmouseover='this.style.textDecoration=" & chr(34)
		       sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
		       sinput = sinput  & "#f7aa22" & chr(34) & ";' xonmouseout='this.style.backgroundColor=" & chr(34)
		       sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
		       sinput = sinput  & "none" & chr(34) & ";' "
		       if session("base")="K"  then
		        sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,"fdesde",GTCSMT,GTCFCH,GTCFCHDAT,"",21,2,0,0,0)
		       else
		       sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,"fdesde",GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
		       end if
		       sinput = sinput  & "  TITLE='Modificar fecha.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"


             sInput = dFechan & "&nbsp;" & sinput & inputFdesde

Celda = Celda & ofv.gencelda("","ttt","","","","","",sInput,"")
Celda = Celda & ofv.gencelda("","ttt","","","","","","Fecha Publicacion Hasta","")

'sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
'sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
'sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

' CELDAS DE HASTA
'if clng(XFSRH) = 0 then
'	sInput2 = "------------"
'else
'	sInput2 = fechaConvert(clng(XFSRH) + clng(cdate("01-01-1900")) )
'	swd = datepart("w",sInput2)
'	sInput2 = Fechador(XFSRH)
'END IF

'sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","no") & ofv.GenerarInput("Hasta3",cdate(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","no")

inputFhasta = ofv.GenerarInput("fhasta",fh,"hidden",sForm,"50","dt","no")

		if isnumeric(fh) then


		      dFecha = cdate(fh + clng(cdate("01-01-1900")))
		      xfecact = fh
		   '   dFechan = Fechador(xfecact)
		      dFechan = GTFechador(fh)

	      else
	          dFecha = ""
			 xfecact = "N"
		      dFechan = "-------"

         end if

		      sImagenI4 = "<B><i class='las la-calendar' style=font-size:14pt;color:" & Session("CMFPNBOT") & "></i></B>"
		      sinput = "<a class=cc style='text-decoration:none;background-color:none;' xonmouseover='this.style.textDecoration=" & chr(34)
		       sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
		       sinput = sinput  & "#f7aa22" & chr(34) & ";' xonmouseout='this.style.backgroundColor=" & chr(34)
		       sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
		       sinput = sinput  & "none" & chr(34) & ";' "
		       if session("base")="K"  then
		        sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,"fhasta",GTCSMT,GTCFCH,GTCFCHDAT,"",21,2,0,0,0)
		       else
		       sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,"fhasta",GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
		       end if
		       sinput = sinput  & "  TITLE='Modificar fecha.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"


             sInput = dFechan & "&nbsp;" & sinput & inputFhasta


Celda = Celda & ofv.gencelda("","ttt","","","","","",sInput,"")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'FILA ESPACIO
Celda = ofv.gencelda("","colspan=4","","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'FILA BOTONES
Celda = ofv.gencelda("","","","","","","","&nbsp;","si")
Celda = Celda & ofv.gencelda("","","center","","","","",btnCons,"si")
Celda = Celda & ofv.gencelda("","","center","","","","",btnLimp,"si")
Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'TABLA CONTENEDORA
tabla = ofv.GenTabla("","aaa","center","80%","","0","","0","0","","",Fila,"si")

'FORM CONSULTA
sAccion = "formEnv.asp"
sHTML = sHTML & ofv.GenForm(sForm,"",sAccion,"","post","",tabla,"si")

'FORM LIMPIAR
sAccion = "formEnv.asp"
xFrmLimp = ofv.GenForm("limpiar","",sAccion,"","post","","","si")

'DIV FORMULARIOS ENCONTRADOS
if acc="BSQ" then

sHTML = sHTML & "<div>"

'ENCABEZADO
Celda = ofv.gencelda("","tt","","5%","","","","*","si")
Celda = Celda & ofv.gencelda("","tt","","35%","","","","Descripcion","si")
Celda = Celda & ofv.gencelda("","tt","","35%","","","","Documento","si")
Celda = Celda & ofv.gencelda("","tt","","25%","","","","Fecha publicacion","si")

Fila = ofv.genrow("","","","","","","",Celda,"si")

'QUERY
sql = "select distinct a.codobjeto,a.descri,ISNULL(e.archivo,'') as archivo,c.fecpublic,a.tipobjeto,e.directorio from objetos a "
sql = sql & " inner join tipobjeto b on (a.tipobjeto=b.tipobjeto and b.tipocpt='K') "
sql = sql & " inner join objprty c on (a.codobjeto=c.codobjeto and a.tipobjeto=c.tipobjeto and c.estado=3) "
sql = sql & " inner join objpropiedades d on (a.codobjeto=d.codobjeto and a.tipobjeto=d.tipobjeto) "
sql = sql & " inner join Objhtm e on (a.codobjeto=e.codobjeto and a.tipobjeto=e.tipobjeto) "
sql = sql & " where a.codobjeto > 0 and e.tipo <> 3 and e.archivo <> ''"

'WHERE CODFORM
scodForm = request.form("cform")
if scodForm & "" = "" then scodForm = request.querystring("cform")
if scodForm & "" <> "" then
	sql = sql & " and d.codigoform='" & scodForm & "'"
end if

'WHERE BANCA
sBanca = request.form("banca")
if sBanca & "" = "" then sBanca = request.querystring("banca")
if sBanca & "" = "" or sBanca & "" = "nsnc" then sBanca = "0"
if sBanca & "" <> "0" then
	sql = sql & " and d.banca='" & sBanca & "'"
end if

'WHERE PRODUCTO
sProducto = request.form("producto")
if sProducto & "" = "" then sProducto = request.querystring("producto")
if sProducto & "" = "" or sProducto & "" = "nsnc" then sProducto = "0"
if sProducto & "" <> "0" then
	sql = sql & " and d.producto='" & sProducto & "'"
end if

'WHERE PROCESO
sProceso = request.form("proceso")
if sProceso & "" = "" then sProceso = request.querystring("proceso")
if sProceso & "" = "" or sProceso & "" = "nsnc" then sProceso = "0"
if sProceso & "" <> "0" then
	sql = sql & " and d.proceso='" & sProceso & "'"
end if

'WHERE DESCRI
sDescri = request.form("descri")
if sDescri & "" = "" then sDescri = request.querystring("descri")
if sDescri & "" <> "" then
	sql = sql & " and a.descri like '%" & sDescri & "%'"
end if

'WHERE FEC PUBLIC DESDE
sFecDesde = request.form("fdesde")
if sFecDesde & "" = "" then sFecDesde = request.querystring("fdesde")
if sFecDesde & "" <> "" then
	'sFecDesde = CLng(CDate(sFecDesde)) - nFecini
	sql = sql & " and c.fecpublic >='" & sFecDesde & "'"
end if

'WHERE FEC PUBLIC HASTA
sFecHasta = request.form("fhasta")
if sFecHasta & "" = "" then sFecHasta = request.querystring("fhasta")
if sFecHasta & "" <> "" then
	'sFecHasta = CLng(CDate(sFecHasta)) - nFecini
	sql = sql & " and c.fecpublic <='" & sFecHasta & "'"
end if

sql = sql & " order by a.descri"
'response.write sql
Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

xTrColor = "#FFFFFF"

Set fso = SERVER.CreateObject("Scripting.FileSystemObject")

do while not rs.eof

fpub = rs(3) + CLng(CDate("01-01-1900"))
fPublic = CDate(fpub)
fPublic = FechadorCor(fPublic,"")

chkbox = ofv.GenerarInput("check" & rs(0),rs(0) & "-" & rs(4),"checkbox","","10","","")

Celda = ofv.gencelda("","ut","","3%","","","",chkbox,"si")
Celda = Celda & ofv.gencelda("","ut","","29%","","","",rs(1),"si")
Celda = Celda & ofv.gencelda("","ut","","28%","","","",rs(2),"si")
Celda = Celda & ofv.gencelda("","ut","","15%","","","",fPublic,"si")

if fso.FileExists(session("DirUpload") & rs("directorio") & "\" & rs("archivo")) then

	Fila = Fila & ofv.genrow("","aaa bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

if xTrColor = "#FFFFFF" then
	xTrColor = "#F0F0F0"
ELSE
	xTrColor = "#FFFFFF"
END IF

end if

rs.movenext
loop

Celda = ofv.gencelda("","a colspan=7","","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

'marcar todo
onclk = "onclick=" & chr(34) & "seleccTodo('" & xForm & "',true);" & chr(34)
btnTodos = ofv.GenerarInput("","Marcar todo","button","","10","bt " & onclk,"")

'desmarcar todo
onclk = "onclick=" & chr(34) & "seleccTodo('" & xForm & "',false);" & chr(34)
btnNing = ofv.GenerarInput("","Desmarcar todo","button","","10","bt " & onclk,"")

Celda = ofv.gencelda("","a colspan=7","center","","","","",btnTodos & btnNing,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","a colspan=7","","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","tt colspan=2","","","","","","Tipo envio","si")

chkboxAdj = "<input id='checkEnv' name='checkEnv' class='dt' type='radio' value='adj'>"

chkboxUrl = "<input id='checkEnv' name='checkEnv' class='dt' class='dt' type='radio' value='url'>"

chkboxCor = "<input id='checkEnv' name='checkEnv' class='dt' class='dt' type='radio' value='cor'>"

boxes = "Adjunto " & chkboxAdj & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Url " & chkboxUrl
boxes = boxes & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CorExpress " & chkboxCor

Celda = Celda & ofv.gencelda("","tt ","","","","","",boxes,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","a colspan=7","","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

onclk = "onclick=" & chr(34) & "document.getElementById('" & xForm & "').submit();" & chr(34)
btnCont = ofv.GenerarInput("","Continuar","button","","10","bt " & onclk,"")
Celda = ofv.gencelda("","a colspan=7","center","","","","",btnCont,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

sTabla = ofv.GenTabla("","aaa","center","100%","","0","","0","0","","",Fila,"si")

if request.form("VolvBotBb") = "no" then

cform = request.querystring("cform")
banca = request.querystring("banca")
producto = request.querystring("producto")
proceso = request.querystring("proceso")
descri = request.querystring("descri")
fdesde = request.querystring("fdesde")
fhasta = request.querystring("fhasta")

sAccion = "formEnvDet.asp?opc=cont&cform=" & cform & "&banca=" & banca & "&producto=" & producto & "&proceso=" & proceso & "&descri=" & descri & "&fdesde=" & fdesde & "&fhasta=" & fhasta
else
sAccion = "formEnvDet.asp?opc=cont&" & request.form
end if
sHTML = sHTML & ofv.GenForm(xForm,"",sAccion,"","post","",sTabla,"si")

sHTML = sHTML & "</div>"

end if


'****** CODIGO PARA CALENDARIO
'sInputDesde=ofv.GenerarInput("Desde",fechaconvert(XFSRD+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
'sAccionRetornoD="formEnv.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
'formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")

'sInputHasta=ofv.GenerarInput("Hasta",fechaconvert(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
'sAccionRetornoH="formEnv.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
'formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta,"si")

' ****** CODIGO PARA CALENDARIO
'tabla= ofv.GenTabla("","aaa bgcolor='#eeeeee'  style='border-style:solid;border-width:1;border-color:#dddddd;' ","","80% ","","0","","5","3","","",formx,"si")
'Formulario = formulario & ofv.GenForm("Subconsulta","",sAccion,STARGET,"post" & funxt,"",TABLA,"si")  & formRetornoD & formRetornoH

Agregar = ""
Buscar = ""
Celda = ""

sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


sHTML = sHTML & "<script type='text/javascript'>"

sHTML = sHTML & "function seleccTodo(fid,dato)"
sHTML = sHTML & " {"
sHTML = sHTML & " var checkboxes = new Array(); "
sHTML = sHTML & " checkboxes = document.getElementById(fid).getElementsByTagName('input'); "
sHTML = sHTML & "for (var i=0; i<checkboxes.length; i++)  { "
sHTML = sHTML & "if (checkboxes[i].type == 'checkbox')   { checkboxes[i].checked=dato;} "
sHTML = sHTML & " }"
sHTML = sHTML & " }"
sHTML = sHTML & "function ejecutar(dato)"
sHTML = sHTML & " { var f=document.getElementById(dato);"
sHTML = sHTML & " f.action=f.action + '?acc=BSQ';"
sHTML = sHTML & " f.submit()"
sHTML = sHTML & " }"
sHTML = sHTML & "</script>"

sHTML = sHTML & xFrmLimp & Formulario & GTCCALENDAR("")& "</body></html>"
end if

Response.Write sHTML

Function Fechador(wkfecha)

fec = cdate(wkfecha + clng(cdate("01-01-1900")))
ndia = ofv.cdia(Weekday(fec))
nMes = ofv.cmes(Month(fec))

Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)

End Function

Function fechaConvert(wkfecha)

Nfecha=cdate(wkfecha)
Nfecha=dateserial(year(wkfecha),month(wkfecha),day(wkfecha))
fechaConvert=Nfecha

end function

%>