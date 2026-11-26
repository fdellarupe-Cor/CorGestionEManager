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
  sTipo = request.querystring("Tipobjeto")
  ID = request.querystring("ID")

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover) & "_ID=" & ID
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2) & "_ID=" & ID
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2) & "_ID=" & ID
  


  response.write ofv.menuheader("#dddddd","")


  set cn = ofv.conectar(ofv.strconn2)


  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "Select * From Objestr Where objid = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bsecuen = ""
  bDescri = " "
  if not rs.eof then
     bsecuen = rs(4)
     bDescri = rs(3)
  end if
  call ofv.cerrarconsulta(rs)

  response.write ofv.GenTabla("","","","100%","","0","","0","0","","","","no")

  bstylo = " style={border-style:ridge;border-width:2;} "

  bicon = "<font face='corflaticon' size=3 color=#ffffff>1</font>"

  Formulario = ofv.GenCelda("","tt" & bstylo,"","90%","","","",bicon & "&nbsp;Detalle de " & bdescri,"si")
    'Test
    sInput = ofv.GenerarInput("","Test","submit","","20","bt","")
    Celda = ofv.GenCelda("","tt" & bstylo,"","","","","",sInput,"si")
    sAccion = "Coestrex.asp?Codobjeto=" & nCodigo & "&Tipobjeto="
    sAccion = sAccion & sTipo & "&sdocum=" & ID 

    Formulario = Formulario & ofv.GenForm("AgreDet1","",sAccion,"_blank","post","",Celda,"si")

    'Volver
    sInput = ofv.GenerarInput("","Volver al Indice","submit","","20","bt","")
    Celda = ofv.GenCelda("","tt" & bstylo,"","","","","",sInput,"si")
    sAccion = "edestr.asp?Codobjeto=" & nCodigo & "&Tipobjeto="
    sAccion = sAccion & sTipo
    Formulario = Formulario & ofv.GenForm("volDet1","",sAccion,"Principal","post","",Celda,"si")
  Fila = ofv.GenRow("","","","","","","",Formulario,"si")
  response.write Fila

  response.write ofv.GenFTabla()


 call ofv.cerrarconsulta(rs)
 call ofv.cerrarconn(cn)
shtml = "</body></html>"

end if

  response.write shtml

%>