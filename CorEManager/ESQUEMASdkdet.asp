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


  nCodigo = "0"
  sTipo   = request.querystring("Tipobjeto")
  sVolver = request.querystring("volver")
  sCod    = "9" 'request.querystring("cod")
  sCIRCUITO    = request.querystring("CIRCUITO")


  vuelta  = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")

  if sCIRCUITO<> "" then
   session("sCIRCUITO")=sCIRCUITO
  else
   sCIRCUITO=session("sCIRCUITO")
  end if


  sHTML = ofv.FormHeader(session("FormN"))



  set cn = ofv.conectar(ofv.strconn2)

 sPerm="COMENTARIOS"


  Celda = ofv.gencelda("","ttt colspan=3","center","2%","","","","AUTORIZADOS&nbsp;PARA&nbsp;" & sPerm,"si")
  Fila = fila & ofv.GenRow("","","","","20","","",Celda,"si")

  Celda = ofv.gencelda("","ttr","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.gencelda("","ttr colspan=2","","2%","","","","Usuario/Grupo","si")
  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")







   IF session("modalidad") = "PRGCOR" THEN

       Strsql = "SELECT OBJID, * FROM ORRSKCIRCUITOSDKDET  "
       Strsql = Strsql & " Where CODRSK=" & sCIRCUITO
       Strsql = Strsql & " And tiposeg = '" & sCod & "'"

      WKSPTABLA = "ORRSKCIRCUITOSDKDET"


   ELSE



       Strsql = "SELECT * FROM OBJSDKdet  "
       Strsql = Strsql & " Where codobjeto=0"
       Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
       ' incidente nro 1374
       Strsql = Strsql & " And tiposeg ='9'"

      WKSPTABLA = "OBJSDKdet"

  END IF
  ' response.write strsql

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof

    sForm = "Sdk" & rs(0)

    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver="
      sHref = sHref  & vuelta & "?Cod=" & sCod & "&CIRCUITO=" & SCIRCUITO
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Permiso",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    sUbic  = instr(rs(4),"USR-")
    sBusca = mid(rs(4),5,len(rs(4))-4)
    sGrpUsr = "S/D"

    if not sUbic > 0 then
       sTabla = "Grupo"

       set cx = ofv.conectar(ofv.strconn0)
       Strsql = "SELECT * FROM " & sTabla & "s"
       Strsql = Strsql & " Where ID = '" & sBusca & "' "
       set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
       if not rx.eof then sGrpUsr = sTabla & ": " & rx(0) & " - " & rx(1)
       call ofv.cerrarconsulta(rx)
       call ofv.cerrarconn(cx)


    else
       sTabla = "Usuario"

        xxxx     = getInfoUser(sBusca)
        sGrpUsr = session("SGU_nombre")

        sGrpUsr = sTabla & ":&nbsp;" & sBusca & " - " & sGrpUsr



    end if





    sinput = ofv.gencelda("","ttt colspan=2","","60%","","","",sGrpUsr,"si")

    Celda = Celda & sinput

    Fila  = fila & ofv.GenRow("","","","","20","","",Celda,"si")

   rs.movenext
   loop

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","0","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"

 if bModificar then
  sAccion = "AgregarPermisoESQUEMA.asp?Codobjeto=" & nCodigo & "&CIRCUITO=" & SCIRCUITO
  sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=" & vuelta
  Agregar = ofv.BotonAgregar(sAccion)
END IF
  sAccion  = "TIPOSDEESQUEMAPROP.asp?" & REQUEST.QUERYSTRING & "&wPerfil=" &sTipo
  Retornar = ofv.BotonVolver(sAccion)





  sHTML = sHTML & "</center>"



'--------------------------------------------------------------------------------
sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>