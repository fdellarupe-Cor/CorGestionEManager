<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  response.redirect "edgeneralbus.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False


  vuelta = ofv.vueltaasp2("../")


  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof Then
    sHTML = ofv.FormBody

    wdescri = rs(3)

    'Descripcion
    Celda = ofv.GenCelda("","tt","","","","","","Descripcion","si")
    sInput = ofv.GenerarInput("",rs(3),"readonly","","90","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    'Identificacion
    Celda = ofv.GenCelda("","tt","","","","","","Identificacion","si")
    sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    'Reemplazo
    Celda = ofv.GenCelda("","tt","","","","","","Reemplaza","si")
 '  if bModificar then
      Strsql = "Select * From csGenObjetos order by xdescri"
      sInput = ofv.GenerarCombo("document.edGeneral",rs(4).name,rs(4).value,Strsql,cn,0,1,"sb")
 '   else
 '     if rs(4) <> "" then
 '     Strsql = "Select xdescri From csGenObjetos2 where codobjeto = " & rs(4)
 '       set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '       if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
 '       call ofv.cerrarconsulta(rsInstrumento)
 '     end if
 '     sInput = ofv.GenerarInput("",sInstrumento,"readonly","","90","dt","")
  '  end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



    Celda = ofv.GenCelda("","aaa colspan=2","","100%","","","","&nbsp;","si")
    Fila =  Fila & ofv.GenRow("","","","","","","",Celda,"si")





    sAccion = ""
 '   if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objetos&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
 '   end if
    Formulario = ofv.GenForm("edGeneral","",sAccion,"","post","",Fila,"si")


  call ofv.cerrarconsulta(rs)




    sHTML = sHTML & "<br><br><center>"
    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

    'Volver




    sInput = "<input  type=SUBMIT  size=10%  value='Volver' class=bt >"
    Celda = ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")
    sAccion = "edgeneral.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    Formulario = ofv.GenForm("edGeneral22","",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")





    sHTML = sHTML & "<center>"
  end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>