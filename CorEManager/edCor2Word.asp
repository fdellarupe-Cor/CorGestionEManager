<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

  XXXX = GENLOGEXT("EDIDET")

  response.expires=0
  server.scripttimeout = 300
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Id")
  sTipo = request.querystring("Tipobj")
  sTitulo = request.querystring("titulo")
  sperfil = request.querystring("perfil")


  sTitulo2 = replace(stitulo,"_"," ")
  shtml=""

ofv.CorFormHeader = CorFormHeader2("Conversion a Word","S")
  response.write ofv.FormHeader("Conversion a Word")

  response.write "<center><br><br><p ><font face=tahoma size=2 color=#993333><b>Conversion de " & stitulo2 & ", Confirma el proceso ?...</b></font></p>"

'<a href=edcor2word.asp?ID=155&Titulo=NORMA:_NG103_-_NORMA_GENERAL&perfil=OBJMAIN&tipobj=NMA Target=_blank><IMG SRC=Imagenes\botrep.jpg title=" Convertir&nbsp;a&nbsp;Word&nbsp;NORMA: NG103 - NORMA GENERAL" border=0></a>

  response.write "<br><br><form action="
  response.write chr(34) & "edcor2wordGen.asp?ID=" & nCodigo & "&Titulo=" & sTitulo & "&perfil=" & sperfil & "&tipobj=" & sTipo & chr(34)
  response.write " method=post >"
  response.write " <table style={border-style:outset;border-width:2;background-color:#f7eedd;} cellspacing=0 cellpadding=3 width=20% ><tr><td align=center ><input type=submit value=Convertir class=bt></td>"
  response.write "<td align=center ><input type=button value=Cancelar class=bt onclick='window.close()' ></td></tr></table>"
  response.write " </form>"

end if

response.write sHTML

%>