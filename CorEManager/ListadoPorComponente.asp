<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else


  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

  'nCodigo = request.querystring("Codcomp")
  sTipo = request.querystring("Tipo")
 ' vuelta = ofv.vueltaasp2("../")

'ofv.CorFormHeader = ofv.menuheader("","")
'sHTML = ofv.FormHeader(session("FormN"))
 response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
  response.write "<tr>"
  response.write "<td width=100% >"
  response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
  response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
  response.write "<tr>"
  response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
  response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
  response.write "</tr></table>"

 stipo=Request.querystring("Tipo")




id = Request.querystring("ID")
set cn = ofv.conectar(ofv.strconn2)
strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 From Complementos,TIPOCOMP "
STRSQL = STRSQL & " Where TIPOCOMP=tipo and tipocomp='" & stipo & "' AND COMPLEMENTOS.COMPUN =  " & SESSION("UN") & "  order by Complementos.descri"
'response.write strsql
set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rs.eof then
sRepHeadv = RepHeader2(rs("descri2") , "600")
else
sRepHeadv = RepHeader2("" , "600")
end if
call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)


shtmlz=generarlistacomponentes2(stipo)
response.write shtml & shtmlz
sPath = Server.mappath(Request.ServerVariables("SCRIPT_NAME"))
sPath2 = Mid(sPath, 1, InStrRev(sPath, "\", -1))

sHeadb = OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")

    Celda = OFV.GenCelda("", "vt", "right", "100%", "", "", "", "PAGINA&nbsp;&&XX de &&XY", "si")
    sHeadb = sHeadb & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


    Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
    sHeadb = sHeadb & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenFTabla()

sheadb=replace(sheadb,"&&","")

end if




Public Function RepHeader2(STitulo, nAncho)

'''        Public RCliLogo
'''        Public RCliNomb
    Dim s, Fecha

    If nAncho = "" Then nAncho = "600"

    s = "<table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
    s = s & "<tr align=top>"
    If RCliLogo <> "" Then s = s & "<td width=4% ><img src=" & RCliLogo & " border=0 ></td>"
    s = s & "<td class=ietx align=left >&nbsp;" & RCliNomb & "</td></tr></table>"
    s = s & "<br><table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
    s = s & "<tr align=top><td class=iett align=left width=70% >" & STitulo & "</td>"
    Fecha = OFV.fechador()
    s = s & "<td class=ietx align=right>" & Fecha & "</td></tr></table>"
    s = s & "<hr color=#000088 size=1 width=" & nAncho & " align=left>"

    RepHeader2 = s

End Function


Function GenerarListaComponentes2(xTipo)


Dim id, tipo, nid, sTipo, sind, i
Dim strsql, simagen, sComplemento, sNombre, ruta, Fila, filax, Celda, sestr, sdeta, sInput
Dim rs2, rs
Dim cn
dim dshtml
    'Response.write "ok"

    dshtml=""
    Set cn = OFV.conectar(OFV.strconn2)



  strsql = "Select descri,archivo From Tipocomp Where tipo = '" & xTipo & "'"

  Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
  If Not rs.EOF Then
    sComplemento = rs(0)
    simagen = rs(1)
  End If
  Call OFV.cerrarconsulta(rs)
  'Response.write strsql


  'strsql = "Select * From Complementos Where Tipocomp='" & xTipo & "'"
  strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 From Complementos,TIPOCOMP "
  STRSQL = STRSQL & " Where TIPOCOMP=tipo and tipocomp='" & xTipo & "' AND COMPLEMENTOS.COMPUN =  " & SESSION("UN") & "  order by Complementos.descri"
 ' Response.write strsql
  Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
  'Response.write strsql
  Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Descripcion", "si") & OFV.GenCelda("", "tt", "", "", "", "", "", "Identificación", "si")
  filax = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
  Fila = ""

  Do While Not rs.EOF
    sNombre = rs(1) & "  -  " & rs(3)
    sNombre = OFV.convertircar(sNombre, " ", "&nbsp;")


    'Descripcion
     Celda = OFV.GenCelda("", "ut", "", "", "", "", "", rs(3), "si")

   'Identificacion
    Celda = Celda & OFV.GenCelda("", "ut ", "", "", "", "", "", rs(1), "si")
    Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
    Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
    Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    rs.movenext
    'exit do

  Loop
   dshtml = dshtml & OFV.GenTabla("", "", "", "80% ", "", "0", "", "0", "0", "", "", filax & Fila, "si")
  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)

  GenerarListaComponentes2 = dshtml




   Call OFV.cerrarconsulta(rs)
   Call OFV.cerrarconn(cn)
End Function




%>

<!-- #INCLUDE FILE="IncFile/parser.inc" -->