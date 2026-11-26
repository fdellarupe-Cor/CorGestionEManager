<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  SFUNCION     = request.querystring("FUNCION")
  nCodigo      = request.querystring("Codcomp")
  sTipo        = request.querystring("Tipocomp")
  sNombre      = request.querystring("Nombre")
  sPerfil      = request.querystring("Perfil")
  sComplemento = request.querystring("Complemento")
  sComplemento = ofv.convertircar(sComplemento,"_"," ")
''  sFuncion     = "E"


  sSClick  = request.querystring("SSClick")
  sNIV1    = request.querystring("NIV1")
  sNIV2    = request.querystring("NIV2")
  sNIV3    = request.querystring("NIV3")
  sNIV4    = request.querystring("NIV4")
  sNIV5    = request.querystring("NIV5")


 if  sSClick = "" then     sSClick      = 99



  set cn = ofv.conectar(ofv.strconn2)
scodi = sCodigo
stipi = sTipo
srestric = "N"

session("NCliente") = sCodigo
session("TCliente") = sTipo

'  XXXX = GENOBJALOG()



  if sFuncion = "E" then
 '    XXXX = GETOBJALOGUPD("ED")
  else
 '    XXXX = GETOBJALOGUPD("CON")
  end if

  xsCorEsq = sNombre


  sCodigo = scodi
  sTipo = stipi

  sHTML = "<html>"


     wxs = "<script language=JavaScript>" & Chr(13)
    ' wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
  '   wxs = wxs & "  if (document.getElementById('ICONACT')) { " & chr(13)
    ''' wxs = wxs & " alert('yaaaa!!'); " & Chr(13)
     wxs = wxs & " document.ICONACT.submit(); " & Chr(13)
  '   wxs = wxs & " } else { " & chr(13)
  '   wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
  '   wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
  '   wxs = wxs & "  } " & chr(13)
     wxs = wxs & " } " & Chr(13)


     wxs = wxs & "function AbrirWin() { " & Chr(13)
  '''   wxs = wxs & " document.festad.submit(); " & Chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs


  if sFuncion = "E" then
     sHTML = sHTML & "<body   onunload='document.init.submit();AbrirWin();' >"
  else
     sHTML = sHTML & "<body   onunload='AbrirWin();' >"
  end if

  sInput = "<a>" & sNombre & ":&nbsp;" & sComplemento & "</a>"
  xsclase = "at"


 ' IF  Session("FULLGUI") = "N" THEN
 '     styl = "background-color:#888888;"
 ' else
 '     styl = XspEx("#888888","#cccccc",0) & ";"
 ' end if

  sHTML = sHTML & OFV.GenTabla("DBpanel", "aaa  ", "", "100%", "", "0", "", "-2", "2", "", "", "", "no")

    h1 = OFV.gencelda("", "at style='padding-left:3px;' ", "", "2", "", "", "", "<IMG valign=middle border=0 src='../coress20/imagenes/logocorPnewicoSm1.gif' style='filter: chroma(color=#ffffff);' >", "si")
    h1 = h1 & ofv.GenCelda("",xsclase,"","","","","",sInput,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    sHTML = sHTML & ofv.GenFTabla()

  if session("ModMain") = "WK" and sFuncion = "E" then
     xstrsol     = 0
  else
     xstrsol     = SSCLICK
  end if

     Session("Segct") = "S"
     Session("Segmo") = "S"
     Session("Segco") = "S"
     Session("Segce") = "S"

''     response.write stipo & " " & ncodigo & " " & sPerfil & " " & sNombre & "<br>"


  if sFuncion = "E" then

    sHTML = sHTML & ofv.GenTabla("","aaa style='display:none;' ","","100%","",0,0,0,0,"","","","no")



    xverasp = left(Session("ASP"),4)
    xverasp = Ucase(xverasp)

    if xverasp = "COMP" then
       xverasp = Session("ASP") & "' target='info"
    else
       xverasp = Session("ASP") & "' target='fesquemas"
    end if

    sHTML = sHTML & ofv.GenFTabla()
    sHTML = sHTML & "<form name=init id=init action='" & xverasp & "' "
    sHTML = sHTML & " method=post style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

  end if

    xverasp = "BarraEnCierra2.asp?"
    xverasp = xverasp & "USR=" & REPLACE(Session("Usuario") & "*"," ","_")
    xverasp = xverasp & "&NETUSR=" & REPLACE(Session("UsuarNT") & "*"," ","_")
    xverasp = xverasp & "&DMN=" & REPLACE(Session("DOMINIO") & "*"," ","_")
    xverasp = xverasp & "&LID=" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_")
    xverasp = xverasp & "&LIDOBJ=" & REPLACE(SESSION("LOGIDDET") & "*"," ","_")

    sHTML = sHTML & "<form name=festad  id=festad action='" & xverasp & "' "
    sHTML = sHTML & " method=post target=Estad style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    xSolasp = ""
    xSolasp = xSolasp & "BarraCompIcon.asp?Tipo=" & sTipo & "&Codigo="
    xSolasp = xSolasp & nCodigo & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
    xSolasp = xSolasp & "&Funcion=" & sFuncion & "&SSCLICK=" & xstrsol & "&Menu=" & sMenu & "&Sector=" & sSector
    xSolasp = xSolasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5



    sHTML = sHTML & "<form name=ICONACT id=ICONACT action='" & xSolasp  & "' "
    sHTML = sHTML & " method=post target='WK2' style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " window.top.CorEsq = '" & xsCorEsq & " / " & sComplemento & "'; " & Chr(13)
     wxs = wxs & " window.top.focus(); " & Chr(13)
     wxs = wxs & " IniWin(); "
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

%>
