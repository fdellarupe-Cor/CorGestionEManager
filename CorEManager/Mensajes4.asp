<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

sMsg = ofv.convertircar(request.querystring("msg"),"_"," ")
sVolver = request.querystring("volver")
'''sHTML = ofv.FormHeader("Atencion")

 tabla = ofv.mensajes(sMsg,sVolver,nBotones)

 tablan = instr(tabla,"&Eliminar=")

 if tablan = 0 then
    tabla = replace(tabla,"asp?","asp?Eliminar=NO&")
 end if

sHTML = sHTML & "</BODY></html>"

response.write sHTML

%>