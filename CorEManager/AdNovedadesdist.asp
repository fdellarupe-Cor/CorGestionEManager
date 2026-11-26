<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  sestado = 1
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
 ' response.write request.querystring
  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar
'  if session("Nomodificar") = "S" then bModificar = False

  CantRegMoverx = 0 'request.querystring("retorno")

  if CantRegMoverx = "" then  CantRegMoverx = request.querystring("Cantregmover")

  ID = request.querystring("ID")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID
'response.write ID & "--"
  sHTML = ofv.FormHeader(session("FormN") & " - Lista de Distribucion del Mensaje")

  set cn = ofv.conectar(ofv.strconn2)
  set cnx = ofv.conectar(ofv.strconn0)

  Celda = ofv.gencelda("","tt colspan=4","center","2%","","","","Mensaje NGP" & id,"si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  Celda = ofv.gencelda("","tt","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.gencelda("","tt colspan=3","","98%","","","","&nbsp;&nbsp;Destinatario","si")

  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  'Strsql = "SELECT * FROM Novegenpardist  a inner join "
  'Strsql = Strsql & " Where a.codigo=" & ID


  Strsql ="SELECT a.*,b.codemp FROM Novegenpardist a"
  Strsql =strsql & " inner join Novegenpar b on (b.objid=a.codigo) "
  Strsql =strsql & " where a.codigo=" & ID


  'response.write (chr(13) & strsql & chr(13))

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then

  do until rs.eof

    sForm = "dist" & rs(0)

    if bModificar then
       sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
       sHref = sHref & "&Tabla=novegenpardist&volver=" & vuelta & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Usuario",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'destinatario

    sTabla = "Usuario"

    sBusca = rs(2)
    codemp="0" & rs("codemp")
    if Ucase(sBusca) = "SECT-" then

       sGrpUsr = "S/D"


       Strsql = "SELECT * FROM eco"
       Strsql = Strsql & " Where eco = 0" & rs(3)

	   'response.write (chr(13) & strsql & chr(13))

       set rx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rx.eof then sGrpUsr = "<font color=#993333><b>Sector: </b></font>" & rx(1)
       call ofv.cerrarconsulta(rx)


       sinput = ofv.gencelda("","vt  colspan=3 style='border-top-color:#dddddd;border-top-style:solid;border-top-width:1;' ","","100%","","","",sGrpUsr,"si")


    elseif Ucase(sBusca) = "TODOS" then

       sGrpUsr = "S/D"


       Strsql = "SELECT * FROM CLIENTEPLAN"
       Strsql = Strsql & " Where ID = 0" & rs(3)
       set rx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rx.eof then sGrpUsr = "<font color=#336699><b>EMPRESA: </b></font>" & rx(1)
       call ofv.cerrarconsulta(rx)


	   'response.write (chr(13) & strsql & chr(13))

       sinput = ofv.gencelda("","vt  colspan=3 style='border-top-color:#dddddd;border-top-style:solid;border-top-width:1;' ","","100%","","","",sGrpUsr,"si")


    else

       sGrpUsr = "S/D"
       set cx = ofv.conectar(ofv.strconn0)

      ' Strsql = "SELECT * FROM " & sTabla & "s"
      ' Strsql = Strsql & " Where ID = '" & sBusca & "' "
      'RESPONSE.WRITE sBusca
      ' sBusca=SPLIT(sBusca,"-")

       XXX=GETINFOUSERex(sBusca,codemp)
      ' set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
       'if not rx.eof then sGrpUsr = "<font color=#005500><b>" & sTabla & ": </b></font>" & rx(0) & " - " & rx(1)
       'sGrpUsr = "<font color=#005500><b>" & sTabla & ": </b></font>" & sBusca(1) & " - " & session("SGU_nombre")
       sGrpUsr = "<font color=#005500><b>" & sTabla & ": </b></font>" & sBusca & " - " & session("SGU_nombre")
       call ofv.cerrarconsulta(rx)
       call ofv.cerrarconn(cx)

       sinput = ofv.gencelda("","vt  colspan=3 style='border-top-color:#dddddd;border-top-style:solid;border-top-width:1;' ","","100%","","","",sGrpUsr,"si")

    end if

    Celda = Celda & sinput

    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

   rs.movenext
   loop

    else


       Celda  = ofv.gencelda("","vt  colspan=4 style='border-top-color:#dddddd;border-top-style:solid;border-top-width:1;' ","","100%","","","","&nbsp;","si")
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    end if

    if bModificar then

       celda = ofv.gencelda("","","","1%","","","","&nbsp;","si")

       AgregarA = "AgregarUsuNoved.asp?ID=" & ID & "&ta=E"
       AgregarA = AgregarA & "&CantRegmover=" & CantRegmoverx & "&volver=" & vuelta
       Agregar = ofv.GenerarInput("bac1","Agregar Usuario","submit","","10","bt","")
       Agregar = ofv.GenCelda("","","right","50%","","","",Agregar,"si")
       Agregar = ofv.GenForm("Agregar","",AgregarA,"","post","",Agregar,"si")

      ' AgregarA = "Agregarnovedist.asp?ID=" & ID & "&ta=S"
      ' AgregarA = AgregarA & "&CantRegmover=" & CantRegmoverx & "&volver=" & vuelta
      ' Agregar2 = ofv.GenerarInput("bac1","Agregar Sector","submit","","10","bt","")
      ' Agregar2 = ofv.GenCelda("","","right","50%","","","",Agregar2,"si")
      ' Agregar2 = ofv.GenForm("Agregar2","",AgregarA,"","post","",Agregar2,"si")

    'IF SESSION("EMPMUN") = "S" then

    '   AgregarA = "Agregarnovedist.asp?ID=" & ID & "&ta=E"
    '   AgregarA = AgregarA & "&CantRegmover=" & CantRegmoverx & "&volver=" & vuelta
    '   Agregar3 = ofv.GenerarInput("bac1","Agregar Empresa","submit","","10","bt","")
    '   Agregar3 = ofv.GenCelda("","","right","50%","","","",Agregar3,"si")
    '   Agregar3 = ofv.GenForm("Agregar2","",AgregarA,"","post","",Agregar3,"si")
    'else
    '   Agregar3 = ofv.GenCelda("","","right","50%","","","","&nbsp;","si")
    'end if

    '   celda = celda & ofv.gencelda("","vt colspan=2 ","","1%","","","",Agregar  & Agregar2  & Agregar3,"si")

       Fila = fila & ofv.GenRow("","","","","","","",celda & Agregar,"si")

   end if

    sInput = ofv.GenerarInput("","Volver","submit","","10","bt","")
    Celda = ofv.GenCelda("","aaa colspan=4","","10%","","","",sInput,"si")

    sAccion = "Adnovedades.asp?CantRegMover=" & CantRegMoverx

    Formulario = ofv.GenForm("Boton" & ID,"",sAccion,"","post","",Celda,"si")

    Fila = fila &  ofv.genrow("","","","","","","",Formulario,"si")

    Celda = ofv.gencelda("","vt","","2%","","","","&nbsp;","si")
    celda = celda & ofv.gencelda("","vt","","80%","","","","&nbsp;","si")
    celda = celda & ofv.gencelda("","vt","","4%","","","","&nbsp;","si")
    celda = celda & ofv.gencelda("","vt","","4%","","","","&nbsp;","si")


    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")




  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aa  border=0","","100%","","0","0","1","0","","",fila,"si")

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","","","")
  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML
%>