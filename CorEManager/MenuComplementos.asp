<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codcomp")
  sTipo = request.querystring("Tipocomp")
  sNombre = request.querystring("Nombre")
  sPerfil = request.querystring("Perfil")
  sComplemento = request.querystring("Complemento")
  sComplemento = ofv.convertircar(sComplemento,"_"," ")

  dim Nombres(6)
  Nombres(1) = "General"
  Nombres(2) = "Propiedades"
  Nombres(3) = "Estructura"
  Nombres(4) = "Detalle"
  Nombres(5) = "Imagenes"
  Nombres(6) = "Compuesto"
  dim ASPs(6)
  ASPs(1) = "edGeneral2.asp"
  ASPs(2) = "edPropiedades2.asp"
  ASPs(3) = "edEstructura2.asp"
  ASPs(4) = ""
  ASPs(5) = "edImagen2.asp"
  ASPs(6) = "edCompuesto2.asp"

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Comperfil Where tipoperfil='" & sPerfil & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  redim Solapas(1)
  redim Solapas2(1)
  Solapas(1) = Nombres(1)
  Solapas2(1) = ASPs(1)
  x = 1
  for z = 2 to 6
    if z <> 4 then
      if rs(z) Then
        x = x + 1    
        redim preserve Solapas(x)
        redim preserve Solapas2(x)
        Solapas(x) = Nombres(z)
        Solapas2(x) = ASPs(z)
      end if
    end if
  next

  call ofv.cerrarconsulta(rs)

  sHTML = ofv.Menuheader("#336699"," background=Imagenes/back2.gif")

  sInput = "<p align=center><font color=#DCDCDC face=tahoma size=1><b>" & sComplemento
  sInput = sInput & "</b></font></p>"
  Celda = ofv.GenCelda("","","","","20","#336699","Imagenes/topban.jpg",sInput,"si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sLink = ofv.GenLink("","","Estructura.asp","","","","ESTRUCTURA","si")
  Celda = ofv.GenCelda("","rr","","","20","#883333","",sLink,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  for x = 1 to ubound(Solapas)
    sSolapa =  ofv.convertircar(solapas(x),"*"," ")
    sHref = solapas2(x) & "?Tipocomp=" & sTipo & "&Codcomp=" & nCodigo
    sLink = ofv.GenLink("","",sHref,"info","","",UCase(sSolapa),"si")
    Celda = ofv.GenCelda("","rr","","","20","#336699","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  next
  sHTML = sHTML & ofv.GenTabla("","","","100% border","","0","","0","0","#cccccc","",Fila,"si")
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>