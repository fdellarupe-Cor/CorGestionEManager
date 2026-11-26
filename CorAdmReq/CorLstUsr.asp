<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sVolver = request.querystring("volver")
  sCod    = request.querystring("cod")
  vuelta  = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")
  PERFIL=request.querystring("PERFIL")
 


  RESPONSE.WRITE REQUEST.QUERYSTRING
  sHTML   = ofv.FormHeader(session("FormN"))

  set cn = ofv.conectar(ofv.strconn6)

  Celda = ofv.gencelda("","ttt colspan=3","center","","","","","<big>Lista De Usuarios </big>","si")
  Fila  = ofv.GenRow("","","","","","","",Celda,"si")

  'Celda = ofv.gencelda("","ttt colspan=3","center","","","","","Lista de Permisos","si")
  'Fila  = fila & ofv.GenRow("","","","","20","","",Celda,"si")

  Celda = ofv.gencelda("","ttr","","2%","","","","&nbsp;","si")
 ' Celda = Celda & ofv.gencelda("","ttr","","1%","","","","Orden","si")
  Celda = Celda & ofv.gencelda("","ttr","","98%","","","","Descripcion","si")
  Fila  = fila & ofv.GenRow("","","","","25","","",Celda,"si")


  xactiva = 0
  xmanda = 0

  Strsql = "SELECT * FROM GPREQXUSR WHERE TIPOREQCOD='" & sCod &"'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  
  
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  Formulario = ""

  vinoSC = 0 

  sImagenU = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=2>&#0128;</FONT></B>"

  do until rs.eof

    sForm = "wfl" & rs(0)

    sHref = ""

    bModificarR = bModificar
    
    xorden = rs(3)
    
    if not isnumeric(xorden) then xorden = 0

    if xmanda and xorden > 8999 then bModificarR = 0

    if bModificarR then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=GPREQXUSR&volver="
      sHref = sHref  & vuelta & "?Cod=" & sCod
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Usuario",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

   
   



    sTabla = "Usuario"
    sBusca = rs(3)
    sGrpUsr = "S/D"

    xxxx     = getInfoUser(sBusca)
    sGrpUsr = session("SGU_nombre")

    sGrpUsr = sImagenU & "&nbsp;" & sBusca & " - " & sGrpUsr



    sinput = ofv.gencelda("","ttt ","","98%","","","",sGrpUsr,"si")

    Celda = Celda & sinput

    Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")


   rs.movenext
   loop

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#999999;' ","","80%","","0","0","-2","1","","",Formulario & fila,"si")
  sHTML = sHTML & "</center>"

  if bModificar then
     sAccion = "AgregarLista.asp?Codobjeto=" & sCod &"&PERFIL=" & PERFIL
     sAccion = sAccion & "&tabla=GPREQXUSR&campo=3&volver=" & vuelta2
     Agregar = ofv.BotonAgregar(sAccion)
  end if
 
  sAccion  = "CORREQPERFILPROPNV.asp?PERFIL=" &  PERFIL
  Retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>