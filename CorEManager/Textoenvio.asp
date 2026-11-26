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
tEnvio = request.form("tEnvio")

if opc = "act" then

tEnvio = request.querystring("tenv")
tTxt = request.form("texto")

sql = "update textoenvio set texto='" & tTxt & "' where tipoenv=" & tEnvio
call ofv.crearconsultaEx(sql,cn,1,parametros)

sHTML = sHTML & "<script>"
sHTML = sHTML & "alert('Texto actualizado');"
sHTML = sHTML & "</script>"


end if

if tEnvio & "" = "" then tEnvio = 1

Celda = ofv.gencelda("","tt","center","15%","","","","Texto para envio:","si")

scomK = 1 & chr(9) & 2 & chr(9) & 3
scomD = "Adjunto" & chr(9) & "Url" & chr(9) & "CorExpress"

combo = ofv.gencomboFX("sForm","tEnvio",tEnvio,scomK,scomD,"si")

sAccion = "TextoEnvio.asp"
sForm = ofv.GenForm("sForm","",sAccion,"","post","",combo,"si")

Celda = Celda & ofv.gencelda("","tt","center","10%","","","",sForm,"si")

Fila = ofv.genrow("","","","","","","",Celda,"si")

sHTML = sHTML & ofv.GenTabla("","aaa","center","25%","","0","","0","0","","",Fila,"si")


sql = "select texto from textoenvio where tipoenv=" & tEnvio
Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)


if not rs.eof then

sInput = "<textarea style='border-radius:5px;border-style:solid;border-width:2px;border-color:" & Session("CMFPNTIT") & ";height:100%;width:100%;' id='texto' name='texto' cols=10 rows=10 >" & rs(0) & "</textarea>"

Celda = ofv.gencelda("","","","","","","",sInput,"")

Fila = ofv.genrow("","","","","","","",Celda,"si")

Celda = ofv.gencelda("","","center","","","","","&nbsp;","si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

btnAct = ofv.GenerarInput("","Actualizar","submit","","10","bt","")
Celda = ofv.gencelda("","","center","","","","",btnAct,"si")
Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

xTabla = ofv.GenTabla("","aaa","center","70%","","0","","0","0","","",Fila,"si")

sAccion = "TextoEnvio.asp?opc=act&tenv=" & tEnvio
sHTML = sHTML & ofv.GenForm("xForm","",sAccion,"","post","",xTabla,"si")

end if







Agregar = ""
Buscar = ""
Celda = ""

sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

'sHTML = sHTML & "</script>"& GTCCALENDAR("")

sHTML = sHTML & xFrmLimp & Formulario & "</body></html>"
end if

Response.Write sHTML

%>