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

ofv.CorFormHeader = ofv.menuheader("","")

  set oPrtComp = server.createobject("CorGeneratorComp.Reportes")
  Set oPrtComp.Request = Request
  Set oPrtComp.Server = Server
  Set oPrtComp.Response = Response
  Set oPrtComp.Session = Session
  Set oPrtComp.Ofv = Ofv
  
  shtmlz=oPrtComp.generarlistacomponentes2()
  'response.write shtmlz 'oPrtComp.shtml
  
 'response.write shtmlz & "222222222"

set oPrt = server.createobject("CorPrintersp.Reportes")

Set oPrt.Request = Request
Set oPrt.Server = Server
Set oPrt.Response = Response
Set oPrt.Session = Session
Set oPrt.Ofv = Ofv

if ofv.explorador = "MSIE" then
   oPrt.AltoPagina = 185
else    
   oPrt.AltoPagina = 130
end if

id = Request.querystring("ID")
set cn = ofv.conectar(ofv.strconn2)
strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 From Complementos,TIPOCOMP "
STRSQL = STRSQL & " Where TIPOCOMP=tipo and tipocomp='" & stipo & "' AND COMPLEMENTOS.COMPUN =  " & SESSION("UN")
'response.write strsql
set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rs.eof then
sRepHead = RepHeader2(rs("descri2") , "600")     
else 
sRepHead = RepHeader2("" , "600")     
end if
call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)
sPath = Server.mappath(Request.ServerVariables("SCRIPT_NAME"))
sPath2 = Mid(sPath, 1, InStrRev(sPath, "\", -1))
sHead = OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")

    Celda = OFV.GenCelda("", "vt", "right", "100%", "", "", "", "PAGINA&nbsp;&&XX de &&XY", "si")
    sHead = sHead & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


    Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
    sHead = sHead & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenFTabla()
 
' response.write cstr(sRepHead)
 sError = oPrt.renderizar(cstr(shtmlz),cstr(shead),cstr(sRepHead),cstr(sPath2))
'response.write oPrt.sPath2 
if sError <> "" then response.write sError

end if

 Function RepHeader2(STitulo, nAncho)

      Dim s, Fecha
    
      If nAncho = "" Then nAncho = "600"
  
      s = "<table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
      s = s & "<tr align=top>"
      If Session("CliLogo") <> "" Then s = s & "<td width=4% ><img src=" & Session("CliLogo") & " border=0 width=22 height=22></td>"
      s = s & "<td class=ut align=left >&nbsp;" & Session("Cliente") & "</td></tr></table>"
      s = s & "<br><table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
      s = s & "<tr align=top><td class=vt align=left width=70% >" & STitulo & "</td>"
      Fecha = OFV.fechador()
      s = s & "<td class=ut align=right>" & Fecha & "</td></tr></table>"
      s = s & "<hr color=#000088 size=1 width=" & nAncho & " align=left>"

      RepHeader2 = s

End Function



%>