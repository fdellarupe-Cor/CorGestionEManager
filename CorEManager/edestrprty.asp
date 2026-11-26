<%

volver = request.querystring("Volver")
nCodigo = request.querystring("Codobjeto")
sTipo = request.querystring("Tipobjeto")
id = request.querystring("ID")
sHTML = "<html><head><title>COR ESolution Suite V4.8 </title></head>"
sHTML = sHTML & "<frameset cols=18%,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Cabecera SCROLLING=auto "
sHTML = sHTML & "src=edestrprtyindice.asp?Volver=" & volver & "&Codobjeto=" & nCodigo
sHTML = sHTML & "&Tipobjeto=" & sTipo & "&ID=" & id & "><FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Contenido "
sHTML = sHTML & "src=nulo.asp></frameset><noframes><body>"
sHTML = sHTML & "<p>Esta página usa marcos, pero su explorador no los admite.</p>"
sHTML = sHTML & "</body></noframes></html>"

response.write sHTML

%>