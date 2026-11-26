<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

  session("xorigen") = "Subtema"

  sTipo = request.querystring("Tipo")
  sNombre = request.querystring("Nombre")
  sNombre = ofv.convertircar(sNombre,"_"," ")
  if request.querystring("Menu") <> "" then Session("Menu") = request.querystring("Menu")

  sID = request.querystring("ID")
  if sNombre = "" Then sNombre = UCase(sTipo)

  sitem = request.querystring("sitem")
  stitem = request.querystring("stitem")
  sticono = request.querystring("sticono")
  stnivel = request.querystring("stnivel")
  Stesq = request.querystring("stesq")
  CantRegMoverxx = request.querystring("CantRegMover")

  xwbody = "bgcolor=#ccbbaa"
  If BrowserName() = "IE" OR BrowserName() = "CHROME" then xwbody = "bgcolor=#cccccc"

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select Prodserv.codprodserv, Prodserv.descri, Prodserv.marcaps, Subtema.descri, "
  Strsql = Strsql & "Subtema.archivo From Prodserv Inner Join Subtema On "
  Strsql = Strsql & "Prodserv.marcaps = Subtema.tipobjeto "
  Strsql = Strsql & "Where Prodserv.codprodserv = " & sID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sTipoE = rs(3)
    sitem = sTipoE & ": " & sitem
    sitem = ofv.ConvertirCar(sitem," ","_")
  end if
  call ofv.cerrarconsulta(rs)

  sHTML = "<html><Title>Normas y Procedimientos</title>" & ofv.FormStyler()
  sHTML = sHTML & "<body " & xwbody & " >"

  Celda = ofv.GenCelda("","mt","center","35%","","","","SUBTEMA","si")
  Fila = ofv.GenRow("","","top","","","","",Celda,"si")
  Tabla = ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")

  If BrowserName() = "IE" OR BrowserName() = "CHROME" then Tabla = Tabla & "<hr color=#000088 size=1>"

  sHref = "Subtemas.asp?CantRegMover=" & CantRegMoverxx
  sImage = ofv.GenImage("","","","","Iconos\ICorMSubtema.jpg","0","")
  sLink = ofv.GenLink("","nt",sHref,"info","","",sImage,"si")
  Celda = ofv.GenCelda("","","","5%","","","",sLink,"si")
  sLink = ofv.GenLink("","nt",sHref,"info","","",ofv.ConvertirCar(sitem,"_"," "),"si")
  Celda = Celda & ofv.GenCelda("","","","","","","",sLink,"si")
  Fila = ofv.GenRow("","","top","","","","",Celda,"si")
  Tabla = Tabla & ofv.GenTabla("","","","","","0","","0","0","","",Fila,"si")

  sImage = ofv.GenImage("","","","","Iconos\ICorMMenu.jpg","0","")

  redim preserve aMenu0(3)
  redim preserve aMenu1(3)
  aMenu0(0) = "TEMAS"
  aMenu1(0) = "TEMAS"
  aMenu0(1) = "DESARROLLO"
  aMenu1(1) = "DESARROLLO"
  aMenu0(2) = "PARA_APROBAR"
  aMenu1(2) = "PARA APROBAR"
  aMenu0(3) = "ELIMINADAS"
  aMenu1(3) = "ELIMINADAS"

  for i = 0 to 3
    sHref = "Objetos.asp?Tipo=ESQUEMAS&ID=" & SID & "&sitem=" & Sitem & "&stitem=" & Stitem
    sHref = sHref & "&sticono=" & Sticono & "&stnivel=" & Stnivel & "&storig=" & aMenu0(i)
    sHref = sHref & "&stesq=" & Stesq
    sLink1 = ofv.GenLink("","nt",sHref,"fesquemas","","",sImage,"si")
    sLink2 = ofv.GenLink("","nt",sHref,"fesquemas","","",aMenu1(i),"si")

    Celda = ofv.GenCelda("","","","2% colspan=2","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","","","5%","","","",sLink1,"si")
    Celda = Celda & ofv.GenCelda("","","","","","","",sLink2,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")
    Tabla = Tabla & ofv.GenTabla("","","","","","0","","0","0","","",Fila,"si")
  next

  if sTipo="Temas" then
    session("sEstado") = "APROBADA"
  elseif sTipo="Desarrollo" then
    session("sEstado") = "DESARROLLO"
  elseif sTipo="ParaAprobar" then
    session("sEstado") = "PARA APROBAR"
  elseif sTipo="Eliminadas" then
    session("sEstado") = "ELIMINADA"
  end if

  sHTML = sHTML & Tabla & "</body></html>"
end if

response.write sHTML

%>