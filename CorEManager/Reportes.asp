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

  Set cn = OFV.conectar(OFV.strconn2)

  strsql = "Select tipobjeto From objetos Where codobjeto = " & nCodigo
  Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

  if not rs.eof then
  sTipo=rs(0)
  end if
  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)
  sTitulo2 = replace(stitulo,"_"," ")
  shtml=""

  sHTML = ofv.FormHeader("Vista Previa de Esquema Normativo")  & chr(13)


    sHTML = sHTML & "<center><br><br><p ><font face=tahoma size=2 color=#993333><b>Vista Previa de " & stitulo2 & ", Confirma el proceso ?...</b></font></p>"


   ' formato="<tr><td align=center ><select name=formato><OPTION value="& chr(34)&"HTML" & chr(34) & ">SALIDA HTML</OPTION><OPTION value="& chr(34)&"WORD"& chr(34)&" >SALIDA WORD</OPTION>"
    if session("APLMDCLI")="BNX" then
   	    formato="<tr><td align=center ><select name=formato><OPTION value="& chr(34)&"WORD" & chr(34) & ">SALIDA WORD</OPTION><OPTION value="& chr(34)&"HTML"& chr(34)&" >SALIDA HTML</OPTION></SELECT> </td><tr>"
   	 else
   	    formato="<tr><td align=center ><select name=formato><OPTION value="& chr(34)&"HTML" & chr(34) & ">SALIDA HTML</OPTION><OPTION value="& chr(34)&"WORD"& chr(34)&" >SALIDA WORD</OPTION></SELECT> </td><tr>"
    end if

    if session("APLMDCLI")="COR" THEN
    formato=formato & "<OPTION value="& chr(34)&"PDF"& chr(34)&" >SALIDA PDF</OPTION>"
    END IF

    formato=formato & "</SELECT> </td><tr>"
    sHTML = sHTML & "<br><br><form action="
    sHTML = sHTML & chr(34) & "reportesgen.asp?ID=" & nCodigo & "&Titulo=" & sTitulo & "&perfil=" & sperfil & "&tipobj=" & sTipo & chr(34)
    sHTML = sHTML & " method=post >"
    sHTML = sHTML & " <table style={border-style:outset;border-width:2;background-color:#f7eedd;} cellspacing=0 cellpadding=3 width=20% >"&formato&"<tr><td align=center ><input type=submit value=Abrir class=bt></td>"
    sHTML = sHTML & "<td align=center ><input type=button value=Cancelar class=bt onclick='window.close()' ></td>"
    sHTML = sHTML & "</tr></table>"
    sHTML = sHTML & " </form>"
  '  sHTML = sHTML & "<A class=aa HREF=PDOC.ASP ><font color=#ff0000>PRUEBA</font></A><br>"
 '    sHTML = sHTML & "<A class=aa HREF=PDOC2.ASP ><font color=#ff0000>PRUEBA2</font></A>"




end if

response.write sHTML

%>