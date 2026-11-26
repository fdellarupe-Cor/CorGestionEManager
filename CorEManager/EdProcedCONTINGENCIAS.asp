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

  nblk = request.querystring("nblk")
  if len(nblk) = 0 then nblk = "D"

  ord = request.querystring("ord")
  if len(ord) = 0 then ord = "T"

  nAccion = request.querystring("Accion")
  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  vueltaC = request.querystring("volver")

  if session("volverproc") = "" then session("volverproc") = vueltaC
  volver2 = ofv.vueltaasp2("../")

  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * From objaccion where CODACCION = '" & nAccion & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     nProced = rs(3)
     nAccCod = rs(18)
     wkactdesc = rs(4)
     wkactsec = rs(10)

        if (rs(4) & "") <> "" then
          Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rs(4)
          set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsAcciones.eof then 
             wkactdesc = rsAcciones(0)
          END IF
          call ofv.cerrarconsulta(rsAcciones)

        end if

  end if

  sHTML = ofv.FormHeader(session("FormN"))

  Strsql = "SELECT * from csprocmaster "
  Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codproced = '" & nProced & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then

      wksecuencia = rs(8) & "." & rs(7)


    'Evento

      sEvento = ""
      if rs(4) <> "" then
        Strsql = "Select descri From Objeven Where objid = " & rs(4)
        set rsEvento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsEvento.eof then sEvento = rsEvento(0)
        call ofv.cerrarconsulta(rsEvento)
      end if

    'Sector

      sSector = ""
      if rs(5) <> "" then

        Strsql = "Select descri From Eco, Objsector where "
        Strsql = Strsql & "Eco.eco = Objsector.codsector and Objsector.objid = " & rs(5)
        set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsSector.eof then sSector = rsSector(0)
        call ofv.cerrarconsulta(rsSector)
      end if


    'Perfil
 
      sPerfil = ""
      if rs(6) <> "" then
        Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(6) & "'"
        set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsPerfil.eof then sPerfil = rsPerfil(0)
        call ofv.cerrarconsulta(rsPerfil)
      end if

    wkprocdesc = wksecuencia & "." & sEvento & "-"  & sSector & "-"  & sPerfil

    Formulario1 = "" 
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Proceso: ","si")
    Celda = Celda & ofv.GenCelda("","ttt","","85%","","","",wkprocdesc,"si")
    Formulario = ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Actividad: ","si")
    Celda = Celda & ofv.GenCelda("","ttt","","85%","","","",wkactsec & "." & wkactdesc,"si")
    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","-1","2","","",Formulario,"si")

  END IF

  Formulario = ""
  xTrColor = "#F9F9F9" 



  Strsql = " Select  a.codcontin,b.codcontin,c.* From PLTIPOCONT c "
  Strsql = Strsql & " left join (select distinct codcontin from plcircalt Where codaccion = " & nacccod & ") a "
  Strsql = Strsql & " on c.tipocontin = a.codcontin "
  Strsql = Strsql & " left join (select distinct codcontin from plcircorr Where codaccion = " & nacccod & ") b "
  Strsql = Strsql & " on c.tipocontin = b.codcontin "
  Strsql = Strsql & " Order By c.TIPOCONTIN "
 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof

      Celda = ofv.GenCelda("","UT","","1%","","","",".","si")


    'Descripcion

    sInput = rs(3)
    Celda = Celda & ofv.GenCelda("","Ut","","90%","","","",sInput,"si")



     'Circuito Alternativo

      sID = "EdProcedCONtiAlt.asp?"
      sID = sID & "Accion=" & nAccCod & "&Codproced=" & nProced  & "&Conti=" & rs(2) & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo

      if (rs(0) & "") = "" then
          xclas  = "btbn"
          sInput = "<font class='btic'  >&#0115;</font>"

      else
          sInput = "<font class='btic'  >&#097;</font>"
          xclas  = "btbs"
      end if


      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Circuito Alternativo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


     'Circuito Correctivo

      sID = "EdProcedCONtiCorr.asp?"
      sID = sID & "Accion=" & nAccCod & "&Codproced=" & nProced  & "&Conti=" & rs(2)  & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo

      if (rs(1) & "") = "" then
          xclas  = "btbn"
          sInput = "<font class='btic'  >&#0115;</font>"

      else
          sInput = "<font class='btic'  >&#097;</font>"
          xclas  = "btbs"
      end if


      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Circuito Correctivo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")





    Fila= ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")

    if xTrColor = "#FFFFFF" then
       xTrColor = "#F9F9F9" 
    ELSE
       xTrColor = "#FFFFFF" 
    END IF



    Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")


    rs.movenext
  loop


  TABLA = ofv.GenTabla("","","","99%","","0","","0","0","","",Formulario,"si")


  CeldaX = ""

  CeldaX = CeldaX & "<div id='DXCONT' style='margin-top:1;margin-left:1;position:absolute;left:5px;"
  CeldaX = CeldaX & "width:95%;height:60%;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
  CeldaX = CeldaX & "background-color:" & session("APLBGD") & ";"
  CeldaX = CeldaX & "' >"

  CeldaX = CeldaX & TABLA

  CeldaX = CeldaX & "</div>"

  sHTML = sHTML & "<CENTER><BR>" & CeldaX & "</CENTER>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""

  'Boton Volver

  sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&XEVEN=" & nProced
  Retornar = ofv.BotonVolver(sAccion)


    'Botonera
    sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",Agregar,Retornar,"")


  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>