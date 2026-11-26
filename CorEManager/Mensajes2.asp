<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

sMsg = ofv.convertircar(request.querystring("msg"),"_"," ")
sVolver = request.querystring("volver")
sID = request.querystring("ID")
nCodigo = request.querystring("Codcomp")
sTipo = request.querystring("Tipocomp")
Tabla = request.querystring("Tabla")
nBotones = CInt(request.querystring("Botones"))

'''sHTML = ofv.FormHeader("Atencion")

if right(sVolver,3) = "asp" then
  sVolver = sVolver & "?Tipocomp=" & sTipo & "&Codcomp=" & nCodigo & "&Tabla=" & Tabla
else
  sVolver = sVolver & "&Tipocomp=" & sTipo & "&Codcomp=" & nCodigo & "&Tabla=" & Tabla
end if

 tabla = ofv.mensajes(sMsg,sVolver,nBotones)

 

 tablan = instr(tabla,"&Eliminar=")

 if tablan = 0 then
    tabla = replace(tabla,"asp?","asp?Eliminar=NO&")
 end if



'----------------------------------------------

  wwshtml = CorPnlFnc("N","")

  wwsHTML = wwsHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

    tabla = replace(tabla,"<TABLE  width=100%", wwsHTML & "<TABLE  width=100%")

sHTML = Tabla

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & "</body></html>"

response.write sHTML

%>